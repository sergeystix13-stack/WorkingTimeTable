// Подсказка для iPhone: как добавить приложение на экран «Домой».
// Показывается только в Safari (не в уже установленном приложении).
(function () {
  "use strict";
  const standalone = window.navigator.standalone === true ||
    (window.matchMedia && window.matchMedia("(display-mode: standalone)").matches);
  const ios = /iPhone|iPad|iPod/.test(navigator.userAgent) ||
    (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  let hidden = false;
  try { hidden = localStorage.getItem("tabel-install-hint") === "hidden"; } catch (e) { /* ок */ }
  if (standalone || hidden) return;
  const box = document.createElement("div");
  box.className = "install";
  box.innerHTML = ios
    ? '<button class="close" aria-label="Скрыть">✕</button><b>Установите Табель на iPhone:</b> нажмите «Поделиться» ' +
      '(квадрат со стрелкой внизу экрана) → «На экран «Домой»» → «Добавить». Дальше открывайте с иконки — ' +
      'данные, введённые здесь в Safari, в приложение не перейдут.'
    : '<button class="close" aria-label="Скрыть">✕</button><b>Можно установить как приложение:</b> в меню браузера ' +
      'выберите «Установить приложение» или «Добавить на главный экран».';
  box.querySelector(".close").addEventListener("click", () => {
    box.remove();
    try { localStorage.setItem("tabel-install-hint", "hidden"); } catch (e) { /* ок */ }
  });
  const today = document.getElementById("today-card");
  today.parentNode.insertBefore(box, today);
})();
