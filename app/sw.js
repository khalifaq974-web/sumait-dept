// عامل خدمة بسيط لتمكين تثبيت التطبيق على الأندرويد (لا يخزّن بيانات الموقع)
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));
self.addEventListener("fetch", () => {});
