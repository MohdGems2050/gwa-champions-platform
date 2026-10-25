import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const src = readFileSync(new URL("../src/App.tsx", import.meta.url), "utf8");
const pick = (re) => { const m = src.match(re); assert.ok(m, `pattern not found: ${re}`); return m; };

/* ── 1. fullscreen is requested only on the application shell ── */
test("fullscreen is requested on the app shell (or documentElement), never on a page/video container", () => {
  const calls = [...src.matchAll(/\.(webkitRequestFullscreen|requestFullscreen|msRequestFullscreen)\(/g)];
  assert.ok(calls.length > 0, "there must be a fullscreen entry point");

  // the target must always come from the SHELL_ID lookup or documentElement
  pick(/const shell=\(document\.getElementById\(SHELL_ID\) as FsElement \| null\) \|\| \(document\.documentElement as FsElement\);/);

  // no element-scoped fullscreen inside video / card / page markup
  assert.ok(!/videoRef\.current\.requestFullscreen/.test(src), "video element must not be fullscreened");
  assert.ok(!/\.requestFullscreen\(\)/.test(src.replace(/shell\.(webkit|ms)?[Rr]equestFullscreen/g, "")),
    "no bare requestFullscreen() on an arbitrary element");
});

/* ── 2. shell wraps the header, so the header stays inside the fullscreen element ── */
test("the shell contains the routed app (header included) and is the fullscreen target", () => {
  pick(/const SHELL_ID = "gwa-app-shell";/);
  pick(/<div id=\{SHELL_ID\}[^>]*>\s*\{children\}/);
  pick(/export default function App\(\)\{\s*return <AppShell><AppRoutes\/><\/AppShell>;/);
});

/* ── 3. native video fullscreen disabled + in-page theater provided ── */
test("videos disable the native fullscreen control and offer an in-page theater instead", () => {
  pick(/controlsList="nodownload nofullscreen"/);
  pick(/function VideoTheater\(/);
  pick(/const \[openTheater,setOpenTheater\]|const\[openTheater,setOpenTheater\]/);
  pick(/function SecureVideo\(props: React\.VideoHTMLAttributes<HTMLVideoElement> & \{ path: string; theater\?: boolean \}\)/);
  assert.ok(src.includes("onDoubleClick={()=>setOpenTheater(true)}"), "double click also opens theater");
});

/* ── 4. theater paints below the header so the header stays visible and usable ── */
test("theater layer sits below the fixed header in z-order and is offset by its height", () => {
  const theater = src.slice(src.indexOf("function VideoTheater("), src.indexOf("/* ═══════ REVIEW / APPROVAL RULES"));
  assert.ok(theater.includes("z-[55]"), "theater must be under the header (z-60)");
  assert.ok(theater.includes("style={{paddingTop:headerHeight}}"), "content must not hide under the header");
  assert.ok(theater.includes("createPortal"), "theater is portalled to the body");
  assert.ok(theater.includes('onKey==="Escape"') || theater.includes('e.key==="Escape"'), "Escape closes theater");
  assert.ok(theater.includes("document.body.style.overflow"), "background scroll is locked while open");
  // every page header still uses the higher layer
  const headers = [...src.matchAll(/<header className="[^"]*z-\[60\][^"]*"/g)];
  assert.ok(headers.length >= 8, `expected the fixed headers to stay above theater, found ${headers.length}`);
});

/* ── 5. global control exists and is reachable inside fullscreen ── */
test("a global fullscreen control is rendered inside the shell", () => {
  const shell = src.slice(src.indexOf("function AppShell("), src.indexOf("export default function App()"));
  assert.ok(shell.includes("useAppFullscreen()"), "shell uses the fullscreen hook");
  assert.ok(shell.includes("onClick={toggle}"), "control toggles fullscreen");
  assert.ok(shell.includes("aria-pressed={isFullscreen}"), "control exposes its state");
  assert.ok(shell.includes("isFullscreen?\"Exit full screen\":\"Full screen\""), "labels reflect state");
});

/* ── 6. header geometry is measured so nothing is obscured ── */
test("header height is measured at runtime and exposed as a CSS variable", () => {
  pick(/const HEADER_VAR = "--gwa-header-h";/);
  pick(/document\.documentElement\.style\.setProperty\(HEADER_VAR/);
  pick(/function useHeaderHeight\(active:boolean\)/);
});

/* ── 7. content and behaviour preserved: videos still stream from the same URLs ── */
test("video sources and page content are untouched by the fix", () => {
  assert.ok(src.includes("const url = useVideoUrl(path);"), "streaming logic unchanged");
  assert.ok(src.includes("<video\n          {...videoProps}\n          src={url}"), "all video props preserved");
  assert.ok(src.includes("path={w.videoUrl}"), "callers unchanged");
  assert.ok(src.includes("return <video {...videoProps} src={url} />;") === false, "video markup was intentionally extended");
  assert.ok(src.includes("{...videoProps}"), "all original video attributes are still forwarded");
});
