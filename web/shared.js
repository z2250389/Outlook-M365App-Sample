(function () {
  const DEFAULT_TARGET_URL = "https://www.ctc-g.co.jp/";
  const DEFAULT_DIALOG_TITLE = "\u30B5\u30A4\u30C9\u30D0\u30FC\u30E9\u30F3\u30C1\u30E3\u30FC";
  const DEFAULT_DIALOG_SIZE = "large";

  function getParam(name) {
    const value = new URLSearchParams(window.location.search).get(name);
    return value == null ? "" : value.trim();
  }

  function getFirstParam(names) {
    for (const name of names) {
      const value = getParam(name);
      if (value) {
        return value;
      }
    }
    return "";
  }

  function toBoolean(value, fallback) {
    if (!value) {
      return fallback;
    }
    const normalized = value.trim().toLowerCase();
    if (["1", "true", "yes", "on"].includes(normalized)) {
      return true;
    }
    if (["0", "false", "no", "off"].includes(normalized)) {
      return false;
    }
    return fallback;
  }

  function sanitizeUrl(value, fallback) {
    const candidate = value || fallback || DEFAULT_TARGET_URL;
    try {
      const url = new URL(candidate, window.location.href);
      if (url.protocol === "http:" || url.protocol === "https:") {
        return url.href;
      }
    } catch (error) {
      // Fall through to fallback below.
    }
    return fallback || DEFAULT_TARGET_URL;
  }

  function parseDialogSize(value, fallback) {
    const candidate = (value || fallback || "large").trim().toLowerCase();
    if (["small", "medium", "large"].includes(candidate)) {
      return candidate;
    }
    return fallback || "large";
  }

  function readConfig() {
    const targetUrl = sanitizeUrl(
      getFirstParam(["TARGET_URL", "targetUrl", "target"]),
      DEFAULT_TARGET_URL
    );
    const dialogTitle =
      getFirstParam(["DIALOG_TITLE", "dialogTitle", "title"]) || DEFAULT_DIALOG_TITLE;
    const autoOpenOnLoad = toBoolean(
      getFirstParam(["AUTO_OPEN_ON_LOAD", "autoOpenOnLoad"]),
      true
    );
    const dialogSize = parseDialogSize(
      getFirstParam(["DIALOG_SIZE", "dialogSize", "size", "DIALOG_WIDTH", "dialogWidth", "width"]),
      DEFAULT_DIALOG_SIZE
    );
    const fallbackUrl = sanitizeUrl(
      getFirstParam(["FALLBACK_URL", "fallbackUrl"]),
      targetUrl
    );

    return {
      targetUrl,
      dialogTitle,
      autoOpenOnLoad,
      dialogSize,
      fallbackUrl,
    };
  }

  function buildDialogUrl(config) {
    const dialogUrl = new URL("dialog.html", window.location.href);
    dialogUrl.searchParams.set("TARGET_URL", config.targetUrl);
    dialogUrl.searchParams.set("DIALOG_TITLE", config.dialogTitle);
    dialogUrl.searchParams.set("FALLBACK_URL", config.fallbackUrl);
    dialogUrl.searchParams.set("DIALOG_SIZE", config.dialogSize);
    return dialogUrl.href;
  }

  function setText(id, value) {
    const node = document.getElementById(id);
    if (node) {
      node.textContent = value;
    }
  }

  function setLink(id, value) {
    const node = document.getElementById(id);
    if (node) {
      node.href = value;
      node.textContent = value;
    }
  }

  function renderConfig(prefix, config) {
    setText(`${prefix}-title`, config.dialogTitle);
    setText(`${prefix}-target`, config.targetUrl);
    setText(`${prefix}-mode`, config.autoOpenOnLoad ? "auto" : "manual");
    setText(`${prefix}-size`, config.dialogSize);
    setLink(`${prefix}-link`, config.targetUrl);
  }

  function writeStatus(prefix, message) {
    setText(`${prefix}-status`, message);
  }

  function supportsDialogUrl() {
    const teams = window.microsoftTeams;
    return Boolean(
      teams &&
        teams.dialog &&
        teams.dialog.url &&
        typeof teams.dialog.url.isSupported === "function" &&
        teams.dialog.url.isSupported()
    );
  }

  function openTargetUrl(targetUrl) {
    window.location.replace(targetUrl);
  }

  async function initializeTeams() {
    const teams = window.microsoftTeams;
    if (teams && teams.app && typeof teams.app.initialize === "function") {
      await teams.app.initialize();
    }
  }

  window.AppShell = {
    readConfig,
    buildDialogUrl,
    renderConfig,
    writeStatus,
    supportsDialogUrl,
    openTargetUrl,
    initializeTeams,
  };
})();
