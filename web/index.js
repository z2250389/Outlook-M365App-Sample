window.addEventListener("DOMContentLoaded", async () => {
  const config = window.AppShell.readConfig();
  const fallbackPanel = document.getElementById("fallback-panel");
  const openButton = document.getElementById("open-button");

  window.AppShell.renderConfig("tab", config);
  window.AppShell.writeStatus("tab", "開いています");

  function showFallback(message) {
    document.body.dataset.state = "fallback";
    fallbackPanel.hidden = false;
    window.AppShell.writeStatus("tab", message);
  }

  try {
    await window.AppShell.initializeTeams();
  } catch (error) {
    // Continue with best effort launch.
  }

  async function launchTarget() {
    const opened = await window.AppShell.openExternalTarget(config.targetUrl);
    if (opened) {
      document.body.dataset.state = "launched";
      return true;
    }

    showFallback("起動できませんでした");
    return false;
  }

  openButton.addEventListener("click", async () => {
    const opened = await launchTarget();
    if (!opened) {
      showFallback("手動で開いてください");
    }
  });

  try {
    if (config.autoOpenOnLoad) {
      const opened = await launchTarget();
      if (!opened) {
        showFallback("手動で開いてください");
      }
    }
  } catch (error) {
    showFallback("起動に失敗しました");
  }
});
