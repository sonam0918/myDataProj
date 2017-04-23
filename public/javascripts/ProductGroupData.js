
var ProductType = require ('./ProductTypeData');

var self = module.exports = {


		 getRevenuebyProdGrp: function (collection, rowNum, ProductGroupSheet, workbook) {
				 	 
         collection.aggregate( [{ $group : { _id : "$product_group", revenue: {$sum : "$billings"} } }],
         function(err,result)
         {

         	// Populate Country Header
			  ProductGroupSheet.set(1, 1, 'Revenue by Product Group');
			  ProductGroupSheet.set(1, 3, 'Product Group');
			  ProductGroupSheet.set(2, 3, 'Revenue');
	 	
         	var i, count;

         	rowNum = 4;
            if(err) throw err;    
            
            for(i=0, count = result.length; i < count; i++) {
            	ProductGroupSheet.set(1, i+rowNum, result[i]._id);
            	ProductGroupSheet.set(2, i+rowNum, result[i].revenue);
            }

				rowNum = rowNum + result.length;
            //console.log(items);
            
            console.log(rowNum);

            self.processProdGrpCC(collection, rowNum, ProductGroupSheet, workbook);

            // Save it 
			workbook.save(function(ok){
			   	if (!ok) 
			      workbook.cancel();
			    else
			      console.log('congratulations, your workbook created');
			});

        });
     },

 processProdGrpCC: function ( collection, rowNum, ProductGroupSheet, workbook) {

     			collection.aggregate( [ {$group: {_id:"$product_group", uniqueCount: {$addToSet : "$base_license"} }  }, { $project: { "country":1, uniqueCustomerCount: {$size:"$uniqueCount"}  } } ]	

					, function(err,result) {
		         	var i, count;

			        rowNum = rowNum + 2; 

				 	ProductGroupSheet.set(1, rowNum, 'Cutomer count per Product Group');    
					ProductGroupSheet.set(1, rowNum + 2, 'Product Group');
			  		ProductGroupSheet.set(2, rowNum + 2, 'Customer Count');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            for(i=0, count = result.length; i < count; i++) {
		            	ProductGroupSheet.set(1, i+rowNum, result[i]._id);
		            	ProductGroupSheet.set(2, i+rowNum, result[i].uniqueCustomerCount);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

					//processCountry();
					self.processAvgTransValuePg(collection, rowNum, ProductGroupSheet, workbook);

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	},

processAvgTransValuePg: function ( collection, rowNum, ProductGroupSheet, workbook) {

     			collection.aggregate( [{ $group : { _id : "$product_group", avg_value: {$avg : "$billings"} } }]	

					, function(err,result) {
		         	var i, count;

			        rowNum = rowNum + 2; 

				 	ProductGroupSheet.set(1, rowNum, 'Average Transaction per Product Group');    
					ProductGroupSheet.set(1, rowNum + 2, 'Product Group');
			  		ProductGroupSheet.set(2, rowNum + 2, 'Average Transaction Value');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            for(i=0, count = result.length; i < count; i++) {
		            	ProductGroupSheet.set(1, i+rowNum, result[i]._id);
		            	ProductGroupSheet.set(2, i+rowNum, result[i].avg_value);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

					//processCountry();
					self.processCummAvgPG(collection, rowNum, ProductGroupSheet, workbook);

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	},


	processCummAvgPG: function ( collection, rowNum, ProductGroupSheet, workbook) {
			collection.aggregate( [{ $group : { _id : "$product_group", avg_value: {$avg : "$billings"} } }]
			 	, function(err,result) {
		         	var i, count, cummValue;

			        rowNum = rowNum + 2; 

				 	ProductGroupSheet.set(1, rowNum, 'Cummulative Average Transaction Value by Product Group');    
					ProductGroupSheet.set(1, rowNum + 2, 'Prodcut Group');
			  		ProductGroupSheet.set(2, rowNum + 2, 'Cummulative Average Transaction Value');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            cummValue = 0;
		            for(i=0, count = result.length; i < count; i++) {
		            	cummValue = cummValue + result[i].avg_value;
		            	ProductGroupSheet.set(1, i+rowNum, result[i]._id);
		            	ProductGroupSheet.set(2, i+rowNum, cummValue);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);


		            var ProductTypeSheet = workbook.createSheet('ProductType', 1000, 1000);
		           	ProductType.getRevenuebyProdTyp(collection, rowNum, ProductTypeSheet , workbook);
		            

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
	 	
