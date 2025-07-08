const debugging = require("office-addin-debugging");

(async () => {
  try {
    await debugging.startDebugging("manifest.xml", {
      app: "word",             // ← specify the Office app
      appType: "desktop",      // ← specify it's desktop (not web)
      devServer: false,        // ← don't launch the dev server
      openDevTools: false      // ← set true if you want dev tools open automatically
    });
  } catch (err) {
    console.error("Sideload failed:", err);
  }
})();