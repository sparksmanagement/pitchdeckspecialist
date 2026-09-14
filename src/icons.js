// Renders react-icons to base64 PNGs for pptxgenjs addImage.
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const fa6 = require("react-icons/fa6");
const md = require("react-icons/md");

const LIB = { ...fa, ...fa6, ...md };

async function icon(name, color = "FFFFFF", size = 256) {
  const Comp = LIB[name];
  if (!Comp) throw new Error(`Unknown icon ${name}`);
  const svg = renderToStaticMarkup(React.createElement(Comp, { color: `#${color}`, size }));
  const png = await sharp(Buffer.from(svg)).resize(size, size).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

module.exports = { icon };
