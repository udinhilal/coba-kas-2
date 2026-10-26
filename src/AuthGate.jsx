import { useEffect, useState } from "react";
import { supabase, supabaseConfigured } from "./lib/supabaseClient";

export default function AuthGate({ children }) {
  const [session, setSession] = useState(null);
  const [ready, setReady] = useState(false);
  const [mode, setMode] = useState("signin");
  const [fullName, setFullName] = useState("");
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [notice, setNotice] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (!supabaseConfigured) { setReady(true); return; }
    let alive = true;
    supabase.auth.getSession().then(({ data, error }) => {
      if (!alive) return;
      if (error) setNotice(error.message);
      setSession(data?.session ?? null);
      setReady(true);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, nextSession) => {
      setSession(nextSession);
      setReady(true);
    });
    return () => { alive = false; subscription.unsubscribe(); };
  }, []);

  const submit = async (event) => {
    event.preventDefault();
    setNotice("");
    if (!supabaseConfigured) {
      setNotice("Supabase belum dikonfigurasi. Isi VITE_SUPABASE_URL dan VITE_SUPABASE_PUBLISHABLE_KEY di environment variables.");
      return;
    }
    if (mode === "signup" && password !== confirmPassword) {
      setNotice("Ulangi kata sandi harus sama.");
      return;
    }
    setBusy(true);
    try {
      if (mode === "signin") {
        const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
        if (error) throw error;
      } else if (mode === "signup") {
        const { data, error } = await supabase.auth.signUp({
          email: email.trim(), password,
          options: { data: { full_name: fullName.trim() }, emailRedirectTo: window.location.origin },
        });
        if (error) throw error;
        if (!data.session) setNotice("Pendaftaran berhasil. Buka email untuk konfirmasi akun, lalu masuk.");
      } else {
        const { error } = await supabase.auth.resetPasswordForEmail(email.trim(), {
          redirectTo: `${window.location.origin}/`,
        });
        if (error) throw error;
        setNotice("Tautan reset kata sandi dikirim ke email.");
      }
    } catch (error) {
      setNotice(error?.message || "Autentikasi gagal. Silakan coba lagi.");
    } finally {
      setBusy(false);
    }
  };

  if (!ready) return <main className="auth-loading">Memeriksa sesi Supabase…</main>;
  if (session) return children({ user: session.user, onSignOut: async () => {
    const { error } = await supabase.auth.signOut();
    if (error) setNotice(error.message);
  } });

  const title = mode === "signin" ? "Masuk ke Kas Warga" : mode === "signup" ? "Daftar akun" : "Lupa kata sandi";
  return <main className="auth-shell"><section className="auth-card">
    <div className="auth-brand"><span className="brand-mark">K</span><div><strong>kas warga</strong><small>RT 26 / RW 04</small></div></div>
    <span className="eyebrow">Supabase Auth</span><h1>{title}</h1>
    <p className="auth-description">{mode === "signin" ? "Sign in untuk masuk ke aplikasi." : mode === "signup" ? "Buat akun pengelola Kas Warga." : "Kami akan mengirim tautan reset ke email."}</p>
    <form className="auth-form" onSubmit={submit}>
      {mode === "signup" && <label>Nama lengkap<input autoComplete="name" value={fullName} onChange={e => setFullName(e.target.value)} required /></label>}
      <label>Email<input type="email" autoComplete="email" value={email} onChange={e => setEmail(e.target.value)} required /></label>
      {mode !== "reset" && <label>Kata sandi<input type="password" minLength={8} autoComplete={mode === "signin" ? "current-password" : "new-password"} value={password} onChange={e => setPassword(e.target.value)} required /></label>}
      {mode === "signup" && <label>Ulangi kata sandi<input type="password" minLength={8} autoComplete="new-password" value={confirmPassword} onChange={e => setConfirmPassword(e.target.value)} required /></label>}
      {notice && <p className="auth-message" role="status">{notice}</p>}
      <button className="primary-btn auth-submit" type="submit" disabled={busy}>{busy ? "Memproses…" : mode === "signin" ? "Sign in" : mode === "signup" ? "Sign up" : "Kirim tautan reset"}</button>
    </form>
    <div className="auth-links">{mode === "signin" ? <><button type="button" onClick={() => { setMode("signup"); setNotice(""); }}>Sign up</button><button type="button" onClick={() => { setMode("reset"); setNotice(""); }}>Lupa password?</button></> : <button type="button" onClick={() => { setMode("signin"); setNotice(""); }}>Kembali ke Sign in</button>}</div>
  </section></main>;
}
