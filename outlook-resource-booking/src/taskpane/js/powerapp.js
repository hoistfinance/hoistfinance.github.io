(function () {
    "use strict";
  
    Office.initialize = function (reason) {
      $(document).ready(function () {

          var appId = "fc8de0d5-b57f-4e2f-8f2d-53ab19be03b6";
        
          $('#canvas-iframe').attr("src", "https://apps.powerapps.com/play/fc8de0d5-b57f-4e2f-8f2d-53ab19be03b6" + appId);
      });
    };
  
  })();
