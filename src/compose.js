import { officeReady, getComposeDays, setComposeDays } from "./lib/office.js";

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
        status(days ? `Follow-up set: ${days} business days.` : "Follow-up is off for this email.");
      } catch (e) {
        status("Couldn't save: " + e.message, true);
      }
    })
  );
});
