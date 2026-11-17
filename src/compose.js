import { officeReady, getComposeDays, setComposeDays, composeMailbox, rememberMailbox } from "./lib/office.js";

const status = (t, err) => {
  const el = document.getElementById("status");
  el.textContent = t;
  el.className = err ? "error" : "muted";
};

officeReady().then(async () => {
  const radios = [...document.querySelectorAll('input[name="days"]')];
  try {
    const days = await getComposeDays();
    radios.forEach((r) => (r.checked = Number(r.value) === days));
    status(days ? `Follow-up set: ${days} business days.` : "Follow-up is off for this email.");
  } catch (e) {
    status("Couldn't read this email's setting: " + e.message, true);
  }
  radios.forEach((r) =>
    r.addEventListener("change", async () => {
      const days = Number(r.value);
      try {
        status("Saving…");
        await setComposeDays(days);
        // Sending from a shared mailbox (e.g. env-doc@)? Make sure "Check follow-ups" looks there too.
        if (days) {
          const mbx = await composeMailbox();
          if (mbx !== "me") await rememberMailbox(mbx);
        }
        status(days ? `Follow-up set: ${days} business days.` : "Follow-up is off for this email.");
      } catch (e) {
        status("Couldn't save: " + e.message, true);
      }
    })
  );
});
