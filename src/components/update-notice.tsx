import { useEffect } from "react";

// Indirizzo di accesso alla versione nuova (Turni 1.0): conta in forma anonima i dispositivi e poi apre Turni 1.0.
const NEW_VERSION_URL = "https://turni-accesso.gabrielecavani.workers.dev";

const STEPS = [
  "Tocca «Apri la nuova versione».",
  "Aggiungi la nuova app alla schermata Home (se usi l’app installata).",
  "Ricarica i file Excel degli orari e scegli di nuovo il tuo nome: i dati di questa versione non vengono trasferiti.",
];

/**
 * Avviso a tutto schermo, senza pulsante di chiusura: invita a passare alla nuova versione (Turni 1.0).
 * Usa solo stili propri (nessuna dipendenza dal resto dell'app), così resta leggibile anche se il resto cambia.
 */
export function UpdateNotice() {
  // Blocca lo scorrimento della pagina sotto l'avviso.
  useEffect(() => {
    const previous = document.body.style.overflow;
    document.body.style.overflow = "hidden";
    return () => {
      document.body.style.overflow = previous;
    };
  }, []);

  return (
    <div
      role="alertdialog"
      aria-modal="true"
      aria-labelledby="update-notice-title"
      aria-describedby="update-notice-text"
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 2147483000,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        padding: "16px",
        paddingTop: "max(16px, env(safe-area-inset-top))",
        paddingBottom: "max(16px, env(safe-area-inset-bottom))",
        background: "rgba(0, 0, 0, 0.6)",
        WebkitBackdropFilter: "blur(10px)",
        backdropFilter: "blur(10px)",
        overflowY: "auto",
      }}
    >
      <div
        style={{
          width: "100%",
          maxWidth: "380px",
          boxSizing: "border-box",
          padding: "20px",
          borderRadius: "28px",
          background: "rgba(255, 255, 255, 0.92)",
          color: "#0f172a",
          border: "1px solid rgba(0, 0, 0, 0.12)",
          boxShadow: "inset 0 1px 0 rgba(255, 255, 255, 0.95), 0 24px 60px rgba(0, 0, 0, 0.25)",
          fontFamily: "-apple-system, system-ui, 'Segoe UI', sans-serif",
        }}
      >
        <h2 id="update-notice-title" style={{ margin: 0, fontSize: "19px", fontWeight: 700, lineHeight: 1.25 }}>
          Questa versione di Calendario Turni non viene più aggiornata
        </h2>
        <p id="update-notice-text" style={{ margin: "10px 0 0", fontSize: "14px", lineHeight: 1.4, color: "#475569" }}>
          È disponibile la nuova versione, Turni 1.0, con le funzioni più recenti e una grafica rinnovata. Questa
          versione sarà rimossa. Per continuare a usare l’app, passa alla nuova versione.
        </p>
        <p style={{ margin: "14px 0 0", fontSize: "12px", fontWeight: 600, color: "#475569" }}>
          Come passare alla nuova versione
        </p>
        <ol style={{ margin: "4px 0 0", padding: 0, listStyle: "none" }}>
          {STEPS.map((text, i) => (
            <li key={text} style={{ display: "flex", gap: "10px", alignItems: "flex-start", marginTop: "8px" }}>
              <span
                aria-hidden="true"
                style={{
                  width: "22px",
                  height: "22px",
                  borderRadius: "11px",
                  background: "rgba(120, 120, 128, 0.14)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: "12px",
                  fontWeight: 700,
                  flexShrink: 0,
                }}
              >
                {i + 1}
              </span>
              <span style={{ fontSize: "14px", lineHeight: 1.35 }}>{text}</span>
            </li>
          ))}
        </ol>
        <a
          href={NEW_VERSION_URL}
          rel="noopener"
          style={{
            marginTop: "18px",
            minHeight: "46px",
            boxSizing: "border-box",
            borderRadius: "23px",
            background: "#003335",
            color: "#ffffff",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: "16px",
            fontWeight: 600,
            textDecoration: "none",
            boxShadow: "inset 0 1px 0 rgba(255, 255, 255, 0.3)",
          }}
        >
          Apri la nuova versione
        </a>
      </div>
    </div>
  );
}
