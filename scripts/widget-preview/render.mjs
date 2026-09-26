// Renders CalendarWidget.json with the sample calendar and saves the Widgets picker screenshots
// (300 x 304 px, medium size, transparent rounded corners) in light and dark themes.
//
//   cd scripts/widget-preview && npm install && npm run render
//
// Uses the installed Microsoft Edge (falls back to Playwright's Chromium), the official Adaptive Cards
// renderer and templating SDK, and a host config tuned to resemble the Windows 11 Widgets Board.
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const require = createRequire(import.meta.url);
const ACData = require('adaptivecards-templating');

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..', '..');
const assets = join(root, 'src', 'OutlookCalendarWidget', 'Assets');
const template = JSON.parse(readFileSync(join(root, 'src', 'OutlookCalendarWidget.Core', 'Templates', 'CalendarWidget.json'), 'utf8'));
const widgetIcon = `data:image/png;base64,${readFileSync(join(assets, 'WidgetIcon.png')).toString('base64')}`;

const size = 'medium';
const data = JSON.parse(execFileSync('dotnet', ['run', join(here, 'SampleCard.cs'), '--', size], { cwd: here, encoding: 'utf8' }));
const card = new ACData.Template(template).expand({ $root: data, $host: { widgetSize: size } });

const themes = {
  light: {
    file: 'WidgetScreenshot.png',
    background: '#FFFFFF',
    border: 'rgba(0, 0, 0, 0.06)',
    text: '#1B1B1B',
    subtle: '#616161',
    accent: '#0F6CBD',
    emphasis: '#08000000',
    button: { background: '#FFFFFF', border: 'rgba(0, 0, 0, 0.14)', text: '#1B1B1B' },
  },
  dark: {
    file: 'WidgetScreenshotDark.png',
    background: '#2B2B2B',
    border: 'rgba(255, 255, 255, 0.08)',
    text: '#FFFFFF',
    subtle: '#C8C8C8',
    accent: '#62ABF5',
    emphasis: '#0FFFFFFF',
    button: { background: '#383838', border: 'rgba(255, 255, 255, 0.1)', text: '#FFFFFF' },
  },
};

function hostConfig(theme) {
  const colors = {
    default: { default: theme.text, subtle: theme.subtle },
    dark: { default: '#1B1B1B', subtle: '#616161' },
    light: { default: '#FFFFFF', subtle: '#C8C8C8' },
    accent: { default: theme.accent, subtle: theme.accent },
    good: { default: '#107C10', subtle: '#107C10' },
    warning: { default: '#BC4B09', subtle: '#BC4B09' },
    attention: { default: '#C50F1F', subtle: '#C50F1F' },
  };
  return {
    fontFamily: "'Segoe UI Variable Text', 'Segoe UI', sans-serif",
    supportsInteractivity: true,
    spacing: { small: 4, default: 8, medium: 12, large: 16, extraLarge: 20, padding: 8 },
    separator: { lineThickness: 1, lineColor: theme.border },
    fontTypes: {
      default: {
        fontFamily: "'Segoe UI Variable Text', 'Segoe UI', sans-serif",
        fontSizes: { small: 12, default: 14, medium: 16, large: 20, extraLarge: 28 },
        fontWeights: { lighter: 300, default: 400, bolder: 600 },
      },
      monospace: {
        fontFamily: "'Cascadia Mono', Consolas, monospace",
        fontSizes: { small: 12, default: 14, medium: 16, large: 20, extraLarge: 28 },
        fontWeights: { lighter: 300, default: 400, bolder: 600 },
      },
    },
    containerStyles: {
      default: { backgroundColor: '#00000000', foregroundColors: colors },
      emphasis: { backgroundColor: theme.emphasis, foregroundColors: colors },
    },
    actions: { maxActions: 5, spacing: 'default', buttonSpacing: 8, actionsOrientation: 'horizontal', actionAlignment: 'left' },
    imageSizes: { small: 40, medium: 80, large: 160 },
  };
}

function page(theme) {
  return `<!doctype html>
<html><head><meta charset="utf-8"><style>
  html, body { margin: 0; background: transparent; }
  .widget { box-sizing: border-box; width: 300px; height: 304px; border-radius: 8px; overflow: hidden;
    background: ${theme.background}; border: 1px solid ${theme.border}; padding: 12px 16px 16px;
    font-family: 'Segoe UI Variable Text', 'Segoe UI', sans-serif; color: ${theme.text}; display: flex; flex-direction: column; }
  .header { display: flex; align-items: center; gap: 8px; height: 24px; margin-bottom: 8px; font-size: 12px; }
  .header img { width: 16px; height: 16px; }
  .header .title { flex: 1; }
  .header .more { font-size: 14px; letter-spacing: 1px; color: ${theme.subtle}; }
  #card { flex: 1; min-height: 0; }
  .ac-adaptiveCard { padding: 0 !important; }
  .ac-pushButton { font: 600 12px 'Segoe UI Variable Text', 'Segoe UI', sans-serif; padding: 4px 12px; border-radius: 4px;
    background: ${theme.button.background}; color: ${theme.button.text}; border: 1px solid ${theme.button.border}; }
  .ac-pushButton.style-positive, .ac-pushButton.primary { background: ${theme.accent}; color: ${theme === themes.dark ? '#000000' : '#FFFFFF'}; border-color: transparent; }
</style></head>
<body><div class="widget">
  <div class="header"><img src="${widgetIcon}" alt=""><span class="title">Outlook Calendar</span><span class="more">···</span></div>
  <div id="card"></div>
</div></body></html>`;
}

let browser;
try {
  browser = await chromium.launch({ channel: 'msedge' });
} catch {
  browser = await chromium.launch();
}

try {
  for (const theme of Object.values(themes)) {
    const tab = await browser.newPage({ viewport: { width: 300, height: 304 }, deviceScaleFactor: 1 });
    await tab.setContent(page(theme));
    await tab.addScriptTag({ path: require.resolve('adaptivecards/dist/adaptivecards.min.js') });
    await tab.evaluate(([payload, config]) => {
      const adaptiveCard = new window.AdaptiveCards.AdaptiveCard();
      adaptiveCard.hostConfig = new window.AdaptiveCards.HostConfig(config);
      adaptiveCard.parse(payload);
      document.getElementById('card').appendChild(adaptiveCard.render());
    }, [card, hostConfig(theme)]);
    await tab.evaluate(() => Promise.all([...document.images].map((img) => img.decode().catch(() => undefined))));
    await tab.evaluate(() => document.fonts.ready);

    const output = join(assets, theme.file);
    await tab.screenshot({ path: output, omitBackground: true, clip: { x: 0, y: 0, width: 300, height: 304 } });
    console.log(`Wrote ${output}`);
    await tab.close();
  }
} finally {
  await browser.close();
}
