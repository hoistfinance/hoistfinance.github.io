(function () {
    "use strict";
  
    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
      $(document).ready(function () {
          var appId = "84ebf9de-0143-47b8-9b30-d662840efb4c";
        
          $('#canvas-iframe').attr("src", "https://apps.powerapps.com/play/" + appId);
      });
    };
  
  })();