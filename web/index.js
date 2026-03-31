window.addEventListener("DOMContentLoaded", async () => {
  const config = window.AppShell.readConfig();
  const dialogUrl = window.AppShell.buildDialogUrl(config);
  const fallbackPanel = document.getElementById("fallback-panel");
  const openButton = document.getElementById("open-button");

  window.AppShell.renderConfig("tab", config);
  window.AppShell.writeStatus("tab", "起動しています");

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

  async function openExternal() {
    const opened = await window.AppShell.openExternalTarget(config.targetUrl);
    if (opened) {
      document.body.dataset.state = "launched";
      return true;
    }

    showFallback("ポップアップがブロックされました");
    return false;
  }

  async function openDialog() {
    if (!window.AppShell.supportsDialogUrl()) {
      return openExternal();
    }

    return new Promise((resolve) => {
      microsoftTeams.dialog.url.open(
        {
          url: dialogUrl,
          title: config.dialogTitle,
          size: config.dialogSize,
          fallbackUrl: config.fallbackUrl,
        },
        (result) => {
          if (result && result.err) {
            openExternal().then(resolve);
            return;
          }

          document.body.dataset.state = "launched";
          resolve(true);
        }
      );
    });
  }

  openButton.addEventListener("click", async () => {
    const opened = await openDialog();
    if (!opened) {
      showFallback("手動で開いてください");
    }
  });

  try {
    if (config.autoOpenOnLoad) {
      const opened = await openDialog();
      if (!opened) {
        showFallback("手動で開いてください");
      }
    }
  } catch (error) {
    showFallback("起動に失敗しました");
  }
});
