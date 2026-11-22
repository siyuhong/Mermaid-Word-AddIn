/* global Office, window */

const OFFICE_JS_URL = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";

let officeJsPromise;

function loadOfficeJs() {
  if (typeof Office !== "undefined" && Office?.onReady) {
    return Promise.resolve();
  }
  if (officeJsPromise) {
    return officeJsPromise;
  }

  // the HTML bootstrap starts this request before the app bundle is parsed.
  if (window.__officeJsReady) {
    officeJsPromise = window.__officeJsReady;
    return officeJsPromise;
  }

  officeJsPromise = new Promise((resolve, reject) => {
    const script = window.document.createElement("script");
    script.src = OFFICE_JS_URL;
    script.async = true;
    script.onload = resolve;
    script.onerror = () => reject(new Error("Could not load Office.js"));
    window.document.head.appendChild(script);
  });
  return officeJsPromise;
}

export async function waitForOfficeReady() {
  await loadOfficeJs();
  if (typeof Office === "undefined" || !Office?.onReady) {
    throw new Error("Office.js did not initialise");
  }
  return Office.onReady();
}
