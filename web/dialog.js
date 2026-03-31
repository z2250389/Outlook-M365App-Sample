window.addEventListener("DOMContentLoaded", async () => {
  const config = window.AppShell.readConfig();
  const frame = document.getElementById("target-frame");
  const fallback = document.getElementById("fallback-link");

  window.AppShell.renderConfig("dialog", config);
  window.AppShell.writeStatus("dialog", "Loading target");

  try {
    await window.AppShell.initializeTeams();
  } catch (error) {
    // Dialog content can still load outside Teams initialization.
  }

  frame.src = config.targetUrl;
  fallback.href = config.fallbackUrl;
  fallback.textContent = config.fallbackUrl;
});
