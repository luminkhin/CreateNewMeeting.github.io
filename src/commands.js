/* global Office */

function setAppointmentTemplate(event) {
  try {
    const item = Office.context.mailbox.item;

    // Ensure compose APIs are available
    if (!item || !item.body || typeof item.body.setAsync !== "function") {
      console.error("Not in appointment compose mode or API unavailable.");
      event.completed();
      return;
    }

    // Subject
    item.subject.setAsync("Project Update Meeting");

    // Location
    item.location.setAsync("Microsoft Teams");

    // Body (HTML)
    const htmlBody = `
      <p>Dear Team,</p>
      <p>This is a scheduled meeting regarding project updates.</p>
      <p><strong>Agenda:</strong></p>
      <ol>
        <li>Progress review</li>
        <li>Issues and blockers</li>
        <li>Next steps</li>
      </ol>
      <p>Best regards,<br/>Minkhin</p>
    `;

    item.body.setAsync(
      htmlBody,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Template inserted successfully.");
        } else {
          console.error("Failed to insert template:", result.error);
        }
        event.completed();
      }
    );
  } catch (e) {
    console.error(e);
    event.completed();
  }
}

// Export for the host
Office.actions.associate("setAppointmentTemplate", setAppointmentTemplate);