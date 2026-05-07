(function () {
  const clock = document.getElementById("bastionClock");
  const fileState = document.getElementById("contextFileState");

  function tick() {
    if (!clock) return;
    const now = new Date();
    clock.textContent = now.toLocaleTimeString("uk-UA", {
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
  }

  tick();
  setInterval(tick, 1000);

  function activateView(view) {
    document.querySelectorAll("[data-view]").forEach((button) => {
      button.classList.toggle("is-active", button.dataset.view === view);
    });

    document.querySelectorAll("[data-view-panel]").forEach((panel) => {
      panel.classList.toggle("is-active", panel.dataset.viewPanel === view);
    });
  }

  document.querySelectorAll("[data-view]").forEach((button) => {
    button.addEventListener("click", () => activateView(button.dataset.view));
  });

  const excelInput = document.getElementById("excelFile");
  if (excelInput && fileState) {
    excelInput.addEventListener("change", (event) => {
      const file = event.target.files?.[0];
      fileState.textContent = file ? file.name : "очікується";
      if (file) activateView("analytics");
    });
  }
})();
