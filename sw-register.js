// Офлайн-режим: регистрация service worker (sw.js). Отдельным файлом — встроенные скрипты запрещены политикой безопасности.
// При каждом запуске проверяем обновление; если пришла новая версия — один раз перезагружаем страницу,
// чтобы сразу показать её (данные сохраняются при каждом изменении, поэтому ничего не теряется).
if ("serviceWorker" in navigator) {
  const hadController = !!navigator.serviceWorker.controller;
  let reloaded = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (!hadController || reloaded) return;   // первая установка — перезагрузка не нужна
    reloaded = true;
    location.reload();
  });
  navigator.serviceWorker.register("sw.js", { updateViaCache: "none" })
    .then(reg => reg.update())
    .catch(() => {});
}
