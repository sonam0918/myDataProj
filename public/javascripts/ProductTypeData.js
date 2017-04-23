
//var ProductType = require ('./ProductTypeData');

var self = module.exports = {


		 getRevenuebyProdTyp: function (collection, rowNum, ProductTypeSheet, workbook) {
				 	 
				 	 console.log('In product type');
         collection.aggregate( [{ $group : { _id : "$product_type", revenue: {$sum : "$billings"} } }],
         function(err,result)
         {

         	// Populate Country Header
			  ProductTypeSheet.set(1, 1, 'Revenue by Product Type');
			  ProductTypeSheet.set(1, 3, 'Product Type');
			  ProductTypeSheet.set(2, 3, 'Revenue');
	 	
         	var i, count;

         	rowNum = 4;
            if(err) throw err;    
            
            for(i=0, count = result.length; i < count; i++) {
            	ProductTypeSheet.set(1, i+rowNum, result[i]._id);
            	ProductTypeSheet.set(2, i+rowNum, result[i].revenue);
            }

				rowNum = rowNum + result.length;
            //console.log(items);
            
            console.log(rowNum);

            self.processProdTypCC(collection, rowNum, ProductTypeSheet, workbook);

            // Save it 
			workbook.save(function(ok){
			   	if (!ok) 
			      workbook.cancel();
			    else
			      console.log('congratulations, your workbook created');
			});

        });
     },

 processProdTypCC: function ( collection, rowNum, ProductTypeSheet, workbook) {

     			collection.aggregate( [ {$group: {_id:"$product_type", uniqueCount: {$addToSet : "$base_license"} }  }, { $project: { "country":1, uniqueCustomerCount: {$size:"$uniqueCount"}  } } ]	

					, function(err,result) {
		         	var i, count;

			        rowNum = rowNum + 2; 

				 	ProductTypeSheet.set(1, rowNum, 'Cutomer count per Product Type');    
					ProductTypeSheet.set(1, rowNum + 2, 'Product Type');
			  		ProductTypeSheet.set(2, rowNum + 2, 'Customer Count');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            for(i=0, count = result.length; i < count; i++) {
		            	ProductTypeSheet.set(1, i+rowNum, result[i]._id);
		            	ProductTypeSheet.set(2, i+rowNum, result[i].uniqueCustomerCount);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

					//processCountry();
					self.processAvgTransValuePt(collection, rowNum, ProductTypeSheet, workbook);

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	},

processAvgTransValuePt: function ( collection, rowNum, ProductTypeSheet, workbook) {

     			collection.aggregate( [{ $group : { _id : "$product_type", avg_value: {$avg : "$billings"} } }]	

					, function(err,result) {
		         	var i, count;

			        rowNum = rowNum + 2; 

				 	ProductTypeSheet.set(1, rowNum, 'Average Transaction per Product Type');    
					ProductTypeSheet.set(1, rowNum + 2, 'Product Type');
			  		ProductTypeSheet.set(2, rowNum + 2, 'Average Transaction Value');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            for(i=0, count = result.length; i < count; i++) {
		            	ProductTypeSheet.set(1, i+rowNum, result[i]._id);
		            	ProductTypeSheet.set(2, i+rowNum, result[i].avg_value);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

					//processCountry();
					self.processCummAvgPt(collection, rowNum, ProductTypeSheet, workbook);

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	},


	processCummAvgPt: function ( collection, rowNum, ProductTypeSheet, workbook) {
			collection.aggregate( [{ $group : { _id : "$product_type", avg_value: {$avg : "$billings"} } }]
			 	, function(err,result) {
		         	var i, count, cummValue;

			        rowNum = rowNum + 2; 

				 	ProductTypeSheet.set(1, rowNum, 'Cummulative Average Transaction Value by Product Type');    
					ProductTypeSheet.set(1, rowNum + 2, 'Prodcut Type');
			  		ProductTypeSheet.set(2, rowNum + 2, 'Cummulative Average Transaction Value');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            cummValue = 0;
		            for(i=0, count = result.length; i < count; i++) {
		            	cummValue = cummValue + result[i].avg_value;
		            	ProductTypeSheet.set(1, i+rowNum, result[i]._id);
		            	ProductTypeSheet.set(2, i+rowNum, cummValue);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

		  		            
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
	 	
