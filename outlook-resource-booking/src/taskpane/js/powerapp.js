(function () {
    "use strict";
  
    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
      $(document).ready(function () {
          var appId = "28d3e0f4-19b1-4f92-ad11-5b7b1dde0841";
        
          $('#canvas-iframe').attr("src", "https://apps.powerapps.com/play/" + appId);
      });
    };
  
  })();