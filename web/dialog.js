window.addEventListener("DOMContentLoaded", async () => {
  const config = window.AppShell.readConfig();
  const openButton = document.getElementById("dialog-open-button");
  const fallback = document.getElementById("fallback-link");

  window.AppShell.renderConfig("dialog", config);
  window.AppShell.writeStatus("dialog", "開いています");

  try {
    await window.AppShell.initializeTeams();
  } catch (error) {
    // Continue with best effort launch.
  }

  fallback.href = config.fallbackUrl;
  fallback.textContent = config.fallbackUrl;
  openButton.addEventListener("click", () => {
    window.location.assign(config.targetUrl);
  });

  window.setTimeout(() => {
    window.location.replace(config.targetUrl);
  }, 80);
});
