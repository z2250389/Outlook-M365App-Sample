(function () {
  const DEFAULT_TARGET_URL = "https://outlook.office.com/mail/";
  const DEFAULT_DIALOG_TITLE = "\u30B5\u30A4\u30C9\u30D0\u30FC\u30E9\u30F3\u30C1\u30E3\u30FC";

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
    const fallbackUrl = sanitizeUrl(
      getFirstParam(["FALLBACK_URL", "fallbackUrl"]),
      targetUrl
    );

    return {
      targetUrl,
      dialogTitle,
      autoOpenOnLoad,
      fallbackUrl,
    };
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
    setLink(`${prefix}-link`, config.targetUrl);
  }

  function writeStatus(prefix, message) {
    setText(`${prefix}-status`, message);
  }

  function openTargetUrl(targetUrl) {
    window.location.replace(targetUrl);
  }

  async function openExternalTarget(targetUrl) {
    const popup = window.open(
      targetUrl,
      "_blank",
      "popup=yes,width=1280,height=900,left=80,top=80,noopener,noreferrer"
    );
    return Boolean(popup);
  }

  async function initializeTeams() {
    const teams = window.microsoftTeams;
    if (teams && teams.app && typeof teams.app.initialize === "function") {
      await teams.app.initialize();
    }
  }

  window.AppShell = {
    readConfig,
    renderConfig,
    writeStatus,
    openTargetUrl,
    openExternalTarget,
    initializeTeams,
  };
})();
