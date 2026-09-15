const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require(process.env.PLAYWRIGHT_PATH || "playwright");

const url = process.env.LEDGERFLOW_URL || "http://127.0.0.1:5173/";
const executablePath = process.env.BROWSER_PATH || undefined;

async function storage(page, key) {
  return page.evaluate((name) => localStorage.getItem(name), key);
}

async function skipTour(page, { waitForGuide = false } = {}) {
  if (waitForGuide) {
    await page.locator("#tourLayer").waitFor({ state: "visible", timeout: 2500 }).catch(() => {});
  }
  const skip = page.locator("#skipTutorialBtn");
  if (await skip.isVisible()) {
    await skip.click();
    await page.locator("#tourLayer").waitFor({ state: "hidden" });
  }
}

async function run() {
  const browser = await chromium.launch({ executablePath, headless: true });
  try {
    const desktop = await browser.newContext({ viewport: { width: 1600, height: 900 }, deviceScaleFactor: 1 });
    const page = await desktop.newPage();
    const errors = [];
    page.on("pageerror", (error) => errors.push(error.message));
    await page.goto(url);
    await page.locator("#homeScreen").waitFor({ state: "visible" });
    assert.equal(await page.locator("#authScreen").isVisible(), false);
    if (process.env.QA_SCREENSHOT_DIR) {
      fs.mkdirSync(process.env.QA_SCREENSHOT_DIR, { recursive: true });
      await page.screenshot({ path: path.join(process.env.QA_SCREENSHOT_DIR, "ledgerflow-home-desktop.png") });
      await page.screenshot({ path: path.join(process.env.QA_SCREENSHOT_DIR, "ledgerflow-home-desktop-full.png"), fullPage: true });
    }
    await page.locator(".home-hero [data-open-trial]").click();
    await page.locator("#trialSampleSelect").selectOption("Utility Store");
    await page.locator("#trialSampleBtn").click();
    await page.locator("#appShell").waitFor({ state: "visible" });
    await page.locator("#tourLayer").waitFor({ state: "visible" });
    await skipTour(page);
    assert.equal(await page.locator("#trialBanner").isVisible(), true);
    assert.equal(await storage(page, "ledgerflow-erp-state-v3"), null);
    const demo = JSON.parse(await storage(page, "ledgerflow-erp-demo-v1"));
    assert.ok(demo.inventory.length > 0, "Sample inventory should be loaded");
    assert.ok(demo.sales.length > 0, "Sample sales should be loaded");
    await page.reload();
    await page.locator("#appShell").waitFor({ state: "visible" });
    await skipTour(page);
    assert.equal(await page.locator("#trialBanner").isVisible(), true);

    const asset = path.join(__dirname, "..", "assets", "ledgerflow-dashboard.png");
    fs.mkdirSync(path.dirname(asset), { recursive: true });
    await page.evaluate(() => {
      document.getElementById("trialBanner").hidden = true;
      document.getElementById("sampleBanner").hidden = true;
      document.getElementById("gettingStartedPanel").hidden = true;
      document.getElementById("exitTrialBtn").hidden = true;
    });
    await page.screenshot({ path: asset });
    console.log(`Captured ${asset}`);
    await page.evaluate(() => (document.getElementById("exitTrialBtn").hidden = false));

    await page.locator("#exitTrialBtn").click();
    await page.locator("#homeScreen").waitFor({ state: "visible" });
    await page.locator(".home-hero [data-open-trial]").click();
    assert.match(await page.locator("#trialBlankTitle").textContent(), /Continue your trial/);
    page.once("dialog", (dialog) => dialog.accept());
    await page.locator("#trialResetBtn").click();
    await page.locator("#appShell").waitFor({ state: "visible" });
    await page.locator("#tourLayer").waitFor({ state: "visible" });
    await skipTour(page);
    const blank = JSON.parse(await storage(page, "ledgerflow-erp-demo-v1"));
    assert.equal(blank.inventory.length, 0);
    assert.equal(blank.sales.length, 0);
    assert.equal(await storage(page, "ledgerflow-erp-state-v3"), null);
    await page.locator(".nav-item[data-view='setup']").click();
    await skipTour(page, { waitForGuide: true });
    await page.locator("#setupTabs [data-tab='setup:fields']").click();
    await skipTour(page, { waitForGuide: true });
    await page.locator("#fieldModuleSelect").selectOption("inventory");
    await page.locator("#fieldLabelInput").fill("Trial shelf label");
    await page.locator("#saveFieldBtn").click();
    const customized = JSON.parse(await storage(page, "ledgerflow-erp-demo-v1"));
    assert.ok(customized.customFields.inventory.some((field) => field.label === "Trial shelf label"));
    await page.locator("#setupTabs [data-tab='setup:business']").click();
    await skipTour(page, { waitForGuide: true });
    await page.locator("#businessNameInput").fill("My Trial Shop");
    await page.locator("#setupForm button[type='submit']").click();
    const updated = JSON.parse(await storage(page, "ledgerflow-erp-demo-v1"));
    assert.equal(updated.organization.name, "My Trial Shop");
    assert.equal(await storage(page, "ledgerflow-erp-state-v3"), null);
    await page.locator("#tutorialBtn").click();
    await page.locator("#tourLayer").waitFor({ state: "visible" });
    await skipTour(page);
    await page.locator("#exitTrialBtn").click();
    await page.locator(".home-header [data-home-auth]").first().click();
    await page.locator("#authScreen").waitFor({ state: "visible" });
    assert.equal(await page.locator("#registerForm").isVisible(), true);
    assert.equal(errors.length, 0, `Browser errors: ${errors.join(" | ")}`);
    await desktop.close();

    const mobile = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 1 });
    const mobilePage = await mobile.newPage();
    const mobileErrors = [];
    mobilePage.on("pageerror", (error) => mobileErrors.push(error.message));
    await mobilePage.goto(url);
    await mobilePage.locator("#homeScreen").waitFor({ state: "visible" });
    if (process.env.QA_SCREENSHOT_DIR) {
      await mobilePage.screenshot({ path: path.join(process.env.QA_SCREENSHOT_DIR, "ledgerflow-home-mobile.png") });
      await mobilePage.screenshot({ path: path.join(process.env.QA_SCREENSHOT_DIR, "ledgerflow-home-mobile-full.png"), fullPage: true });
    }
    await mobilePage.locator("#homeMenuBtn").click();
    assert.equal(await mobilePage.locator("#homeNav").isVisible(), true);
    await mobilePage.locator("#homeNav a[href='#pricing']").click();
    assert.equal(await mobilePage.locator("#homeNav").isVisible(), false);
    const overflow = await mobilePage.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    assert.ok(overflow <= 1, `Mobile horizontal overflow: ${overflow}px`);
    const heroGap = await mobilePage.evaluate(() =>
      document.querySelector(".home-hero-image").getBoundingClientRect().top -
      document.querySelector(".home-trust-line").getBoundingClientRect().bottom,
    );
    assert.ok(heroGap >= 8, `Mobile hero text and image gap: ${heroGap}px`);
    await mobilePage.locator(".home-hero [data-open-trial]").click();
    assert.equal(await mobilePage.locator("#trialDialogBackdrop").isVisible(), true);
    await mobilePage.keyboard.press("Escape");
    assert.equal(await mobilePage.locator("#trialDialogBackdrop").isVisible(), false);
    assert.equal(mobileErrors.length, 0, `Mobile browser errors: ${mobileErrors.join(" | ")}`);
    await mobile.close();
    console.log("Homepage and trial flows passed on desktop and mobile.");
  } finally {
    await browser.close();
  }
}

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
