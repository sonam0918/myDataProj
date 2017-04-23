
var ProductGroup = require ('./ProductGroupData');

var self = module.exports = {


		 processCountry: function (collection, rowNum, CountrySheet, workbook) {
				 	 
         collection.aggregate( [{ $group : { _id : "$country", revenue: {$sum : "$billings"} } }],
         function(err,result)
         {

         	// Populate Country Header
			  CountrySheet.set(1, 1, 'Revenue by Country');
			  CountrySheet.set(1, 3, 'Country');
			  CountrySheet.set(2, 3, 'Revenue');
	 	
         	var i, count;

         	rowNum = 4;
            if(err) throw err;    
            
            for(i=0, count = result.length; i < count; i++) {
            	CountrySheet.set(1, i+rowNum, result[i]._id);
            	CountrySheet.set(2, i+rowNum, result[i].revenue);
            }

				rowNum = rowNum + result.length;
            //console.log(items);
            
            console.log(rowNum);

            self.processCountryCC(collection, rowNum, CountrySheet, workbook);

            // Save it 
			workbook.save(function(ok){
			   	if (!ok) 
			      workbook.cancel();
			    else
			      console.log('congratulations, your workbook created');
			});

        });
     },

 processCountryCC: function ( collection, rowNum, CountrySheet, workbook) {

     			collection.aggregate( [ {$group: {_id:"$country", uniqueCount: {$addToSet : "$base_license"} }  }, { $project: { "country":1, uniqueCustomerCount: {$size:"$uniqueCount"}  } } ]	

					, function(err,result) {
		         	var i, count;

			        rowNum = rowNum + 2; 

				 	CountrySheet.set(1, rowNum, 'Cutomer count per country');    
					CountrySheet.set(1, rowNum + 2, 'Country');
			  		CountrySheet.set(2, rowNum + 2, 'Customer Count');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            for(i=0, count = result.length; i < count; i++) {
		            	CountrySheet.set(1, i+rowNum, result[i]._id);
		            	CountrySheet.set(2, i+rowNum, result[i].uniqueCustomerCount);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

					//processCountry();
					self.processAvgTransValue(collection, rowNum, CountrySheet, workbook);

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	},

processAvgTransValue: function ( collection, rowNum, CountrySheet, workbook) {

     			collection.aggregate( [{ $group : { _id : "$country", avg_value: {$avg : "$billings"} } }]	

					, function(err,result) {
		         	var i, count;

			        rowNum = rowNum + 2; 

				 	CountrySheet.set(1, rowNum, 'Average Transaction per country');    
					CountrySheet.set(1, rowNum + 2, 'Country');
			  		CountrySheet.set(2, rowNum + 2, 'Average Transaction Value');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            for(i=0, count = result.length; i < count; i++) {
		            	CountrySheet.set(1, i+rowNum, result[i]._id);
		            	CountrySheet.set(2, i+rowNum, result[i].avg_value);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

					//processCountry();
					self.processCummAvg(collection, rowNum, CountrySheet, workbook);

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	},


	processCummAvg: function ( collection, rowNum, CountrySheet, workbook) {
			collection.aggregate( [{ $group : { _id : "$country", avg_value: {$avg : "$billings"} } }]
			 	, function(err,result) {
		         	var i, count, cummValue;

			        rowNum = rowNum + 2; 

				 	CountrySheet.set(1, rowNum, 'Cummulative Average Transaction Value by Country');    
					CountrySheet.set(1, rowNum + 2, 'Country');
			  		CountrySheet.set(2, rowNum + 2, 'Cummulative Average Transaction Value');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            cummValue = 0;
		            for(i=0, count = result.length; i < count; i++) {
		            	cummValue = cummValue + result[i].avg_value;
		            	CountrySheet.set(1, i+rowNum, result[i]._id);
		            	CountrySheet.set(2, i+rowNum, cummValue);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

		            var ProductGroupSheet = workbook.createSheet('ProductGroup', 1000, 1000);
		            ProductGroup.getRevenuebyProdGrp(collection, rowNum, ProductGroupSheet , workbook);
		            

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});


	}






} // module exports
	 	
