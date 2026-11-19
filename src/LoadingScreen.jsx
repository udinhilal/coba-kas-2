
export default function LoadingScreen({
  message = "Menyiapkan data kas Anda..."
}) {
  return (
    <main
      className="auth-loading"
      role="status"
      aria-live="polite"
    >
      {/* Animasi buku kas */}
      <div className="loading-visual" aria-hidden="true">

        <div className="loading-orbit loading-orbit-one"></div>
        <div className="loading-orbit loading-orbit-two"></div>

        <div className="loading-book">
          <span className="loading-page loading-page-left"></span>
          <span className="loading-page loading-page-right"></span>
          <span className="loading-book-spine"></span>
        </div>

      </div>

      <h1 className="loading-brand">
        Kas Warga
      </h1>

      <p className="loading-title">
        {message}
      </p>

      {/* Garis animasi pemuatan */}
      <div className="loading-progress">
        <span></span>
      </div>

      <small className="loading-footnote">
        Mohon tunggu sebentar
      </small>

    </main>
  );
}
