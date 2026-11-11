
import { useEffect, useState } from "react";
import { supabase, supabaseConfigured } from "./lib/supabaseClient";

export default function AuthGate({ children }) {
  const [session, setSession] = useState(null);
  const [ready, setReady] = useState(false);
  const [mode, setMode] = useState("signin");
  const [fullName, setFullName] = useState("");
  const [username, setUsername] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [showConfirmPassword, setShowConfirmPassword] = useState(false);
  const [notice, setNotice] = useState("");
  const [busy, setBusy] = useState(false);
  const [communityIdentity, setCommunityIdentity] = useState({
    rtNumber: "26",
    rwNumber: "04",
    kelurahan: "",
    kecamatan: "",
    kabupatenKota: "",
  });

  // Identitas dibaca sebelum login agar halaman masuk selalu memakai nama lingkungan terkini.
  useEffect(() => {
    if (!supabaseConfigured || !supabase) return undefined;
    let alive = true;
    supabase
      .from("community_identity")
      .select("rt_number, rw_number, kelurahan, kecamatan, kabupaten_kota")
      .eq("id", 1)
      .maybeSingle()
      .then(({ data, error }) => {
        if (!alive) return;
        if (error) {
          console.warn("Identitas lingkungan belum dapat dimuat:", error.message);
          return;
        }
        if (data) {
          setCommunityIdentity({
            rtNumber: String(data.rt_number || "26"),
            rwNumber: String(data.rw_number || "04"),
            kelurahan: String(data.kelurahan || ""),
            kecamatan: String(data.kecamatan || ""),
            kabupatenKota: String(data.kabupaten_kota || ""),
          });
        }
      })
      .catch((error) => console.warn("Gagal memuat identitas lingkungan:", error));
    return () => { alive = false; };
  }, []);

  const communityIdentityLine = `RT ${communityIdentity.rtNumber} / RW ${communityIdentity.rwNumber}`;
  const communityLocationLine = [
    communityIdentity.kelurahan ? `Kel. ${communityIdentity.kelurahan}` : "",
    communityIdentity.kecamatan ? `Kec. ${communityIdentity.kecamatan}` : "",
    communityIdentity.kabupatenKota || "",
  ].filter(Boolean).join(" · ");

  useEffect(() => {
    if (!supabaseConfigured) {
      setReady(true);
      return;
    }

    let alive = true;

    supabase.auth.getSession().then(({ data, error }) => {
      if (!alive) return;

      if (error) setNotice(error.message);

      setSession(data?.session ?? null);
      setReady(true);
    });

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event, nextSession) => {
      if (!alive) return;

      setSession(nextSession);
      setReady(true);

      if (event === "PASSWORD_RECOVERY") {
        setMode("update");
        setPassword("");
        setConfirmPassword("");
        setNotice("Silakan buat kata sandi baru.");
      }
    });

    return () => {
      alive = false;
      subscription.unsubscribe();
    };
  }, []);

  const submit = async (event) => {
    event.preventDefault();
    setNotice("");

    if (!supabaseConfigured) {
      setNotice(
        "Supabase belum dikonfigurasi. Periksa environment variables."
      );
      return;
    }

    if (
      (mode === "signup" || mode === "update") &&
      password !== confirmPassword
    ) {
      setNotice("Ulangi kata sandi harus sama.");
      return;
    }

    if (
      (mode === "signup" || mode === "update") &&
      password.length < 8
    ) {
      setNotice("Kata sandi minimal 8 karakter.");
      return;
    }

    setBusy(true);

    try {
      if (mode === "signin") {
        const { data, error } =
          await supabase.auth.signInWithPassword({
            email: email.trim(),
            password,
          });

        if (error) throw error;

        setSession(data.session);
        setEmail("");
        setPassword("");
        setConfirmPassword("");

      } else if (mode === "signup") {
        const { data, error } = await supabase.auth.signUp({
          email: email.trim(),
          password,
          options: {
  data: {
    full_name: fullName.trim(),
    username: username.trim().toLowerCase(),
  },
  emailRedirectTo: window.location.origin,
},
        });

        if (error) throw error;

        if (!data.session) {
          setNotice(
            "Pendaftaran berhasil. Buka email untuk konfirmasi akun, lalu masuk."
          );
        }

      } else if (mode === "reset") {
        const { error } =
          await supabase.auth.resetPasswordForEmail(
            email.trim(),
            {
              redirectTo: window.location.origin,
            }
          );

        if (error) throw error;

        setNotice(
          "Tautan reset kata sandi berhasil dikirim. Silakan periksa email Anda."
        );

      } else if (mode === "update") {
        const { error } = await supabase.auth.updateUser({
          password,
        });

        if (error) throw error;

        // Keluar agar pengguna login kembali
        // menggunakan kata sandi yang baru.
        await supabase.auth.signOut();

        setSession(null);
        setMode("signin");
        setPassword("");
        setConfirmPassword("");
        setShowPassword(false);
        setShowConfirmPassword(false);

        setNotice(
          "Kata sandi berhasil diperbarui. Silakan login menggunakan kata sandi baru."
        );
      }

    } catch (error) {
      setNotice(
        error?.message || "Autentikasi gagal. Silakan coba lagi."
      );
    } finally {
      setBusy(false);
    }
  };

  if (!ready) {
    return (
      <main className="auth-loading">
        Memeriksa sesi Supabase…
      </main>
    );
  }

  // Masuk ke aplikasi jika sesi aktif,
  // kecuali sedang melakukan pemulihan kata sandi.
  if (session && mode !== "update") {
    return children({
      user: session.user,
      onSignOut: async () => {
        const { error } = await supabase.auth.signOut();
        if (error) setNotice(error.message);
      },
    });
  }

  const title =
    mode === "signin"
      ? "Masuk ke Kas Warga"
      : mode === "signup"
      ? "Daftar akun"
      : mode === "reset"
      ? "Lupa kata sandi"
      : "Buat Kata Sandi Baru";

  const description =
    mode === "signin"
      ? "Sign in untuk masuk ke aplikasi."
      : mode === "signup"
      ? "Buat akun pengelola Kas Warga."
      : mode === "reset"
      ? "Masukkan email terdaftar untuk menerima tautan reset."
      : "Masukkan kata sandi baru untuk akun Anda.";

  return (
    <main className="auth-shell">
      <section className="auth-card">

        <div className="auth-brand">
          <span className="brand-mark">K</span>
          <div className="auth-community-meta">
            <strong>kas warga</strong>
            <small className="auth-community-rt">{communityIdentityLine}</small>
            {communityLocationLine && (
              <small className="auth-community-location">{communityLocationLine}</small>
            )}
          </div>
        </div>

        <h1>{title}</h1>

        <p className="auth-description">
          {description}
        </p>

        <form className="auth-form" onSubmit={submit}>

          {mode === "signup" && (
  <>
    <label>
      Nama lengkap
      <input
        autoComplete="name"
        value={fullName}
        onChange={(e) => setFullName(e.target.value)}
        required
      />
    </label>

    <label>
      Username
      <input
        type="text"
        autoComplete="username"
        placeholder="Contoh: hilaludin"
        value={username}
        onChange={(e) =>
          setUsername(
            e.target.value.toLowerCase().replace(/\s/g, "")
          )
        }
        minLength={3}
        maxLength={30}
        pattern="[a-z0-9._-]+"
        title="Gunakan huruf kecil, angka, titik, garis bawah, atau tanda hubung."
        required
      />
      <small>Username digunakan untuk identitas akun dan nama tampilan.</small>
    </label>
  </>
)}
          {mode !== "update" && (
            <label>
              Email
              <input
                type="email"
                autoComplete="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
              />
            </label>
          )}

          {mode !== "reset" && (
            <label>
              Kata sandi

              <div className="auth-password-field">
                <input
                  type={showPassword ? "text" : "password"}
                  minLength={8}
                  autoComplete={
                    mode === "signin"
                      ? "current-password"
                      : "new-password"
                  }
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  required
                />

                <button
                  className="auth-password-toggle"
                  type="button"
                  onClick={() =>
                    setShowPassword((value) => !value)
                  }
                  aria-label={
                    showPassword
                      ? "Sembunyikan kata sandi"
                      : "Tampilkan kata sandi"
                  }
                >
                  {showPassword ? "Sembunyikan" : "Tampilkan"}
                </button>
              </div>
            </label>
          )}

          {(mode === "signup" || mode === "update") && (
            <label>
              Ulangi kata sandi

              <div className="auth-password-field">
                <input
                  type={showConfirmPassword ? "text" : "password"}
                  minLength={8}
                  autoComplete="new-password"
                  value={confirmPassword}
                  onChange={(e) =>
                    setConfirmPassword(e.target.value)
                  }
                  required
                />

                <button
                  className="auth-password-toggle"
                  type="button"
                  onClick={() =>
                    setShowConfirmPassword((value) => !value)
                  }
                  aria-label={
                    showConfirmPassword
                      ? "Sembunyikan ulangi kata sandi"
                      : "Tampilkan ulangi kata sandi"
                  }
                >
                  {showConfirmPassword
                    ? "Sembunyikan"
                    : "Tampilkan"}
                </button>
              </div>
            </label>
          )}

          {notice && (
            <p className="auth-message" role="status">
              {notice}
            </p>
          )}

          <button
            className="primary-btn auth-submit"
            type="submit"
            disabled={busy}
          >
            {busy
              ? "Memproses…"
              : mode === "signin"
              ? "Sign in"
              : mode === "signup"
              ? "Sign up"
              : mode === "reset"
              ? "Kirim tautan reset"
              : "Simpan Kata Sandi Baru"}
          </button>

        </form>

        <div className="auth-links">
          {mode === "signin" ? (
            <>
              <button
                type="button"
                onClick={() => {
                  setMode("signup");
                  setNotice("");
                }}
              >
                Sign up
              </button>

              <button
                type="button"
                onClick={() => {
                  setMode("reset");
                  setNotice("");
                }}
              >
                Lupa password?
              </button>
            </>
          ) : mode !== "update" ? (
            <button
              type="button"
              onClick={() => {
                setMode("signin");
                setNotice("");
                setPassword("");
                setConfirmPassword("");
              }}
            >
              Kembali ke Sign in
            </button>
          ) : null}
        </div>

      </section>
    </main>
  );
}
