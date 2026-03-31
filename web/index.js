window.addEventListener("DOMContentLoaded", async () => {
  const config = window.AppShell.readConfig();

  window.AppShell.renderConfig("tab", config);
  window.AppShell.writeStatus("tab", "Launching");

  try {
    await window.AppShell.initializeTeams();
    window.AppShell.writeStatus("tab", "Ready");
  } catch (error) {
    window.AppShell.writeStatus("tab", "Continue without Teams initialization");
  }

  const openButton = document.getElementById("open-button");
  const manualRow = document.getElementById("manual-row");

  async function openTarget() {
    window.AppShell.writeStatus("tab", "Opening Outlook on the web");
    const opened = await window.AppShell.openExternalTarget(config.targetUrl);
    if (opened) {
      window.AppShell.writeStatus("tab", "Opened");
      return;
    }

    window.AppShell.writeStatus("tab", "Popup blocked. Use the button");
    manualRow.hidden = false;
  }

  openButton.addEventListener("click", async () => {
    const opened = await window.AppShell.openExternalTarget(config.targetUrl);
    window.AppShell.writeStatus("tab", opened ? "Opened" : "Blocked");
  });

  if (config.autoOpenOnLoad) {
    await openTarget();
  } else {
    manualRow.hidden = false;
  }
});
