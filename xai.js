/* Conexión del Campo Gemelo Digital Ganadero con api.x.ai.
   La clave no se guarda en el archivo ni en el informe: solo en esta visita del navegador.
   Las cuentas del lote las hace el gemelo. Esta conexión redacta lo que pidan las solapas nuevas. */
(function () {
  var BASE = "https://api.x.ai/v1";
  var MODELO = "grok-4.7";
  var CLAVE = "gemelo-xai-clave-sesion";
  var SISTEMA = "Sos el redactor del Campo Gemelo Digital Ganadero de la Sociedad Rural de San Luis. Hablás en español de la Argentina, corto y concreto, para un productor ganadero. Las cuentas de carga, hectáreas por equivalente vaca, destete y agua las hace el gemelo: no las recalcules ni las cambies. No propongas recortar el agua de bebida, desmontar ni sembrar soja. Si un dato no viene en el mensaje, decí que falta y no lo inventes. Respondé solo lo que pide la solapa.";

  function leerClave() {
    try { return sessionStorage.getItem(CLAVE) || ""; } catch (e) { return ""; }
  }

  function guardarClave(clave) {
    var limpia = String(clave || "").trim();
    try {
      if (limpia) sessionStorage.setItem(CLAVE, limpia);
      else sessionStorage.removeItem(CLAVE);
    } catch (e) {}
    return limpia;
  }

  function textoDe(data) {
    var out = data && data.output;
    var partes = [];
    if (!Array.isArray(out)) return "";
    out.forEach(function (item) {
      if (!item || item.type !== "message" || !Array.isArray(item.content)) return;
      item.content.forEach(function (c) {
        if (c && c.text && (c.type === "output_text" || c.type === "text")) partes.push(c.text);
      });
    });
    return partes.join("\n").trim();
  }

  function consultar(opts) {
    opts = opts || {};
    var clave = leerClave();
    if (!clave) {
      return Promise.reject(new Error("Falta la clave de api.x.ai. Se carga en Conexión para solapas nuevas."));
    }
    var datos = opts.datos == null ? "" : (typeof opts.datos === "string" ? opts.datos : JSON.stringify(opts.datos));
    var cuerpo = opts.pregunta || "";
    var fotos = [];
    if (opts.imagen) fotos.push(opts.imagen);
    if (Array.isArray(opts.imagenes)) {
      opts.imagenes.forEach(function (img) { if (img) fotos.push(img); });
    }
    if (!fotos.length) {
      cuerpo = "Solapa: " + (opts.solapa || "solapa") + "\nPregunta: " + cuerpo;
      if (datos) cuerpo += "\nDatos del gemelo (no los cambies):\n" + datos;
    }
    var usuario = cuerpo;
    if (fotos.length) {
      usuario = [{ type: "input_text", text: cuerpo || "Leé la imagen." }];
      fotos.forEach(function (img) {
        usuario.push({ type: "input_image", image_url: img, detail: "high" });
      });
    }
    return fetch(BASE + "/responses", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + clave
      },
      body: JSON.stringify({
        model: MODELO,
        store: false,
        max_output_tokens: opts.maxTokens || (fotos.length ? 900 : 600),
        reasoning: { effort: "low" },
        input: [
          { role: "system", content: opts.sistema || SISTEMA },
          { role: "user", content: usuario }
        ]
      })
    }).then(function (r) {
      return r.json().then(function (data) {
        if (!r.ok) {
          var msg = data && data.error && (data.error.message || data.error);
          throw new Error(typeof msg === "string" ? msg : ("La API respondió " + r.status));
        }
        var texto = textoDe(data);
        if (!texto) throw new Error("La API respondió sin texto.");
        return { texto: texto, id: data.id || "", modelo: data.model || MODELO };
      });
    });
  }

  function probar() {
    return consultar({ solapa: "conexión", pregunta: "Respondé solo: conexión lista." });
  }

  window.GemeloXai = {
    base: BASE,
    modelo: MODELO,
    conectada: function () { return !!leerClave(); },
    guardarClave: guardarClave,
    consultar: consultar,
    probar: probar
  };

  function pintarEstado(el, texto, bien) {
    if (!el) return;
    el.hidden = false;
    el.className = bien ? "ok" : "aviso";
    el.textContent = texto;
  }

  function iniciar() {
    var input = document.getElementById("xaiClave");
    var btn = document.getElementById("xaiProbar");
    var estado = document.getElementById("xaiEstado");
    if (!input || !btn) return;
    if (leerClave()) {
      input.placeholder = "Clave cargada en esta visita";
      pintarEstado(estado, "Conectado a api.x.ai en esta visita. Las solapas nuevas ya pueden usarla.", true);
    }
    btn.addEventListener("click", function () {
      var clave = (input.value || "").trim() || leerClave();
      if (!clave) {
        pintarEstado(estado, "Pegá la clave de api.x.ai. No queda en el archivo del gemelo.", false);
        return;
      }
      guardarClave(clave);
      input.value = "";
      btn.disabled = true;
      pintarEstado(estado, "Probando la conexión…", true);
      probar().then(function () {
        pintarEstado(estado, "Conexión lista. Modelo " + MODELO + ". Sirve para las solapas que se sumen.", true);
      }).catch(function (err) {
        pintarEstado(estado, err && err.message ? err.message : "No se pudo conectar.", false);
      }).then(function () {
        btn.disabled = false;
      });
    });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", iniciar);
  else iniciar();
})();
