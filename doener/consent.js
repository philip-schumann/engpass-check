/* Cookie-Banner + Meta-Pixel. Der Pixel lädt erst nach Zustimmung. */
(function () {
  var PIXEL_ID = "1916075619356217";
  var KEY = "sc_consent_v1";
  function get() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
  function set(v) { try { localStorage.setItem(KEY, v); } catch (e) {} }

  function loadPixel() {
    if (window.fbq) return;
    !function(f,b,e,v,n,t,s){if(f.fbq)return;n=f.fbq=function(){n.callMethod?
    n.callMethod.apply(n,arguments):n.queue.push(arguments)};if(!f._fbq)f._fbq=n;
    n.push=n;n.loaded=!0;n.version='2.0';n.queue=[];t=b.createElement(e);t.async=!0;
    t.src=v;s=b.getElementsByTagName(e)[0];s.parentNode.insertBefore(t,s)}(window,
    document,'script','https://connect.facebook.net/en_US/fbevents.js');
    fbq('init', PIXEL_ID);
    fbq('track', 'PageView');
    (window.SC_EVENTS || []).forEach(function (ev) { fbq('track', ev[0], ev[1] || {}); });
  }

  // Seiten melden Events hierüber; sie werden nur mit Zustimmung gesendet.
  window.scTrack = function (name, params) {
    if (get() === "yes") { loadPixel(); fbq('track', name, params || {}); }
    else { (window.SC_EVENTS = window.SC_EVENTS || []).push([name, params]); }
  };

  function banner() {
    var d = document.createElement("div");
    d.id = "cookie";
    d.innerHTML =
      '<p><b>Cookies & Meta-Pixel</b><br>Ich nutze den Meta-Pixel, um zu messen, ob meine Anzeigen funktionieren, und um dir passende Inhalte zu zeigen. Das passiert nur, wenn du zustimmst. Mehr in der <a href="datenschutz.html" target="_blank" rel="noopener">Datenschutzerklärung</a>.</p>' +
      '<div><button type="button" data-v="no">Nur notwendige</button><button type="button" data-v="yes">Akzeptieren</button></div>';
    d.addEventListener("click", function (e) {
      var v = e.target.getAttribute && e.target.getAttribute("data-v");
      if (!v) return;
      set(v); d.remove();
      if (v === "yes") loadPixel();
    });
    document.body.appendChild(d);
  }

  window.scResetConsent = function () { try { localStorage.removeItem(KEY); } catch (e) {} banner(); };

  document.addEventListener("DOMContentLoaded", function () {
    var c = get();
    if (c === "yes") loadPixel();
    else if (c !== "no") banner();
  });
})();
