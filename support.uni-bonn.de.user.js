// ==UserScript==
// @name        Znuny: Bump shown tickets, add "to English" button
// @namespace   github.com/olifre/userstyles
// @match       https://support.uni-bonn.de/*
// @updateURL   https://raw.githubusercontent.com/olifre/userscripts/main/support.uni-bonn.de.user.js
// @version     1.0.0
// @grant       none
// @description Allows to select a larger number of tickets to show, and translate replies to English.
// @author      Oliver Freyermuth <o.freyermuth@googlemail.com> (https://olifre.github.io/)
// @license     Unlicense
// ==/UserScript==

if (/\bAction=AgentDashboard\b/.test (location.search) ) {
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
}


if (/\bAction=AgentTicketEmail\b/.test (location.search) ) {
  var tlBtn=document.createElement('a');
  tlBtn.innerHTML="to English";
  tlBtn.href='#';
  tlBtn.onclick = function() {
   var replyIFrame = document.querySelector('.cke_wysiwyg_frame');
   var replyDocument = replyIFrame.contentDocument || replyIFrame.contentWindow.document;
   var replyBody = replyDocument.body;
   var lines = replyBody.innerHTML.split('<br>');
   for (var l=0; l<Math.min(lines.length,4); ++l) {
     lines[l]=lines[l].replace("Hallo ", "Dear ");
     lines[l]=lines[l].replace("Viele Grüße", "Cheers");
   }
   replyBody.innerHTML = lines.join("<br>");
  };
  var contentDiv = document.querySelector('label[for="RichText"]').parentElement;
  contentDiv.insertBefore(tlBtn, contentDiv.firstChild);
}
