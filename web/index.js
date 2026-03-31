window.addEventListener("DOMContentLoaded", async () => {
  const config = window.AppShell.readConfig();
  const dialogUrl = window.AppShell.buildDialogUrl(config);
  const teams = window.microsoftTeams;
  const fallbackWindowMs = 2000;

  window.AppShell.renderConfig("tab", config);
  window.AppShell.writeStatus("tab", "Initializing");

  try {
    await window.AppShell.initializeTeams();
    if (!window.AppShell.supportsDialogUrl()) {
        window.AppShell.writeStatus("tab", "Dialog API unavailable. Redirecting");
      window.AppShell.openTargetUrl(config.targetUrl);
      return;
    }

    const openDialog = () => {
      const openedAt = Date.now();
      window.AppShell.writeStatus("tab", "Opening dialog");
      teams.dialog.url.open(
        {
          url: dialogUrl,
          title: config.dialogTitle,
          size: config.dialogSize,
          fallbackUrl: config.fallbackUrl,
        },
        (response) => {
          if (response && response.err && Date.now() - openedAt < fallbackWindowMs) {
            window.AppShell.writeStatus("tab", "Dialog failed. Redirecting");
            window.AppShell.openTargetUrl(config.targetUrl);
          }
        }
      );
    };

    if (config.autoOpenOnLoad) {
      openDialog();
    } else {
      window.AppShell.writeStatus("tab", "Ready");
      document.getElementById("open-button").addEventListener("click", openDialog);
      document.getElementById("manual-row").hidden = false;
    }
  } catch (error) {
    window.AppShell.writeStatus("tab", "Initialization failed. Redirecting");
    window.AppShell.openTargetUrl(config.targetUrl);
  }
});
