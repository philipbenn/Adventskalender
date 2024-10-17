'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.sass.setConfig({
  dropCssFiles: true,
  useCSSModules: true,
  warnOnNonCSSModules: false,
  cleanCssOptions: {
    level: 0,
    compatibility: {
      colors: {
        hexAlpha: false, // controls 4- and 8-character hex color support
        opacity: true // controls `rgba()` / `hsla()` color support
      }
    }
    , returnPromise: true
  },
  autoprefixerOptions: { overrideBrowserslist: ["> 1%", "last 2 versions", "not dead"] }
});

/* fast-serve */
const { addFastServe } = require("spfx-fast-serve-helpers");
addFastServe(build);
/* end of fast-serve */

build.initialize(require('gulp'));