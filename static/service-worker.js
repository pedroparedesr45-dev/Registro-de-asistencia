// Service worker mínimo: solo habilita la instalación de la PWA.
// No cachea nada (la app siempre necesita conexión para marcar
// asistencia con GPS y foto en tiempo real).

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  self.clients.claim();
});

self.addEventListener('fetch', () => {
  // Sin caché: siempre va a la red. Esto es intencional para que
  // los trabajadores nunca marquen con datos desactualizados.
});
