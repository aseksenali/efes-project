/* Regenerates the Font Awesome 6 icon PNGs embedded in the deck.
   Output: <repo>/assets/icons/*.png  (256 px, transparent background)

   Usage:  npm install && node gen.js
   Only needs re-running if transform.py starts using a new icon. */
const path = require("path");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa6");

const OUT = path.join(__dirname, "..", "..", "..", "assets", "icons");

// [file name, react-icons component, hex colour]
const ICONS = [
  ["search",        FA.FaMagnifyingGlass,  "5CA3DB"], // AI search & knowledge base
  ["invoice",       FA.FaFileCircleCheck,  "E0A64B"], // 1C invoice validation
  ["calc",          FA.FaCalculator,       "3FB572"], // scaling of calculations
  ["calendar",      FA.FaRegCalendarDays,  "5CA3DB"], // historical promo calendar
  ["quote",         FA.FaQuoteLeft,        "2D83C4"], // pull quotes
  ["bolt_white",    FA.FaBolt,             "FFFFFF"], // accent band
  ["clock_green",   FA.FaRegClock,         "3FB572"], // saved-time chip (light)
  ["clock_dkgreen", FA.FaRegClock,         "15924F"], // saved-time hero
];

(async () => {
  for (const [name, Comp, color] of ICONS) {
    const svg = RDS.renderToStaticMarkup(
      React.createElement(Comp, { color: "#" + color, size: "256" })
    );
    const file = path.join(OUT, name + ".png");
    await sharp(Buffer.from(svg)).png().toFile(file);
    console.log("ok", file);
  }
})();
