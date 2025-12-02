/* global Office */

Office.onReady(() => {
  const subjectEl = document.getElementById("subject");
  const locationEl = document.getElementById("location");
  const bodyEl = document.getElementById("body");
  const applyBtn = document.getElementById("apply");

  applyBtn.addEventListener("click", () => {
    const item = Office.context.mailbox.item;

    if (!item || !item.body || typeof item.body.setAsync !== "function") {
      console.error("Not in appointment compose mode or API unavailable.");
      return;
    }

    // Set subject and location
    item.subject.setAsync(subjectEl.value || "");
    item.location.setAsync(locationEl.value || "");

    // Set body as HTML
    item.body.setAsync(
      bodyEl.value || "",
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Template applied.");
        } else {
          console.error("Failed to apply template:", result.error);
        }
      }
    );
  });
});