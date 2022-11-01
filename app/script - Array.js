'use strict';

// Wrap everything in an anonymous function to avoid poluting the global namespace
(function () {
  // Event handlers for filter change
  let unregisterHandlerFunctions = [];

  let worksheet1, worksheet2;
  // Use the jQuery document ready signal to know when everything has been initialized
  $(document).ready(function () 
  {
	  console.log("Test_4");
  
	   tableau.extensions.initializeAsync().then(function () {
      // Get worksheets from tableau dashboard
	  worksheet1 = tableau.extensions.dashboardContent.dashboard.worksheets[0];
      worksheet1.getSummaryDataAsync().then( function (mydata){
		  
		  console.log(mydata);
		  alert(mydata.data[2][1].value);
		  alert(mydata.data[2][2].value);
		  alert(mydata.data[2][3].value);
	  
	});
	console.log("end");
	});
    
   
  });

})();