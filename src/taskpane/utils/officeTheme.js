/* global Office, window, document */
import { webDarkTheme, webLightTheme } from "@fluentui/react-components";

const DARK_LUMINANCE_THRESHOLD = 0.5;
const THEME_CHECK_INTERVAL_MS = 1000;

/**
 * Convert an Office theme colour (`#RRGGBB`, or `rgb(r,g,b)` in some hosts)
 * into an approximate 0..1 luminance.
 */
export function getThemeLuminance(color) {
  if (!color) {
    return null;
  }

  const value = String(color).trim();
  let red;
  let green;
  let blue;

  const hex = value.startsWith("#") ? value.slice(1) : value;
  if (/^[0-9a-f]{6}$/i.test(hex)) {
    red = parseInt(hex.slice(0, 2), 16);
    green = parseInt(hex.slice(2, 4), 16);
    blue = parseInt(hex.slice(4, 6), 16);
  } else {
    const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(value);
    if (!rgb) {
      return null;
    }
    [, red, green, blue] = rgb.map(Number);
  }

  if (![red, green, blue].every((channel) => Number.isFinite(channel))) {
    return null;
  }

  return (0.299 * red + 0.587 * green + 0.114 * blue) / 255;
}

export function isOfficeThemeDark(color) {
  const luminance = getThemeLuminance(color);
  return luminance !== null && luminance < DARK_LUMINANCE_THRESHOLD;
}

export function getOfficeTheme() {
  return isOfficeDark() ? webDarkTheme : webLightTheme;
}

export function isOfficeDark() {
  const theme = typeof Office === "undefined" ? undefined : Office.context?.officeTheme;
  return typeof theme?.isDarkTheme === "boolean"
    ? theme.isDarkTheme
    : isOfficeThemeDark(theme?.bodyBackgroundColor);
}

/**
 * Word exposes theme values but no OfficeThemeChanged event (that event is
 * Outlook-only). Re-read them while visible and when returning to the pane.
 */
export function subscribeOfficeThemeChange(handler) {
  if (typeof window === "undefined" || typeof document === "undefined") {
    return () => {};
  }

  let isDark = isOfficeDark();
  const checkTheme = () => {
    if (document.hidden) {
      return;
    }

    const nextIsDark = isOfficeDark();
    if (nextIsDark !== isDark) {
      isDark = nextIsDark;
      handler();
    }
  };

  window.addEventListener("focus", checkTheme);
  document.addEventListener("visibilitychange", checkTheme);
  const interval = window.setInterval(checkTheme, THEME_CHECK_INTERVAL_MS);

  return () => {
    window.clearInterval(interval);
    window.removeEventListener("focus", checkTheme);
    document.removeEventListener("visibilitychange", checkTheme);
  };
}
