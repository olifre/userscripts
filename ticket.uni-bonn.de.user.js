// ==UserScript==
// @name        KIX/OTRS: Bump shown tickets
// @namespace   github.com/olifre/userstyles
// @match       https://ticket.uni-bonn.de/*Action=AgentDashboard
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/ticket.uni-bonn.de.user.js
// @version     1.0.0
// @grant       none
// @description Allows to select a larger number of tickets to show.
// @author      Oliver Freyermuth <o.freyermuth@googlemail.com> (https://olifre.github.io/)
// @license     Unlicense
// ==/UserScript==

allConfigSelects = Array.from(document.querySelectorAll('[id^="UserDashboardPref"][id$="-Shown"]'));

allConfigSelects.forEach(
  function (select) {
    allOpts=Array.from(select.options);
    maxCnt=allOpts.reduce((max,val) => {return Math.max(max,val.value)}, 0);
    //console.log(maxCnt);
    allOpts.forEach(
      function (option) {
        if (option.value == maxCnt)
          {
           //console.log("Replacing...");
           option.value="100";
           option.text="100";
          }
      }
    )
  }
);
