var express = require('express');
var path = require('path');
var favicon = require('serve-favicon');
var logger = require('morgan');
var cookieParser = require('cookie-parser');
var bodyParser = require('body-parser');
var MongoClient = require('mongodb').MongoClient;
var excelbuilder = require('msexcel-builder');

var index = require('./routes/index');
var users = require('./routes/users');
var CountryData = require('./public/javascripts/CountryData');

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'jade');

// uncomment after placing your favicon in /public
//app.use(favicon(path.join(__dirname, 'public', 'favicon.ico')));
app.use(logger('dev'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/', index);
app.use('/users', users);

// Connect to the db
MongoClient.connect("mongodb://localhost:27017/myDB", function (err, db) {

	 // Create a new workbook file in current working-path 
	 var workbook = excelbuilder.createWorkbook('./output/', 'myData.xlsx');
	 var rowNum = 0;
	  
	 // Create a new worksheet with 10 columns and 12 rows 
	 var PurchaseYear = workbook.createSheet('PurchaseYear', 1000, 1000);
	 var CountrySheet = workbook.createSheet('Country', 1000, 1000);


     db.collection('transData', function (err, collection) {

     	function processPurchaseYear() {
	     	// Fill some data 
			  PurchaseYear.set(1, 1, 'Revenue by Purchase Year');
			  PurchaseYear.set(1, 3, 'Purchase Year');
			  PurchaseYear.set(2, 3, 'Revenue');
			  rowNum = 4;
			  //for (var i = 2; i < 5; i++)
			  //  PurchaseYear.set(i, 1, 'test'+i);
			  
			 collection.aggregate( [{ $group : { _id : "$purchase_year", revenue: {$sum : "$billings"} } }]
			 	, function(err,result) {
		         	var i, count;
		            if(err) throw err;    
		            
		            for(i=0, count = result.length; i < count; i++) {
		            	PurchaseYear.set(1, i+rowNum, result[i]._id);
		            	PurchaseYear.set(2, i+rowNum, result[i].revenue);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

				 	getCustomerCount();

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});     		
     	}

     	function getCustomerCount() {

		 	collection.aggregate( [ { $group: {_id:"$purchase_year", uniqueCount: {$addToSet : "$base_license"} }  }, { $project: { "Purchase_Year":1, uniqueCustomerCount: {$size:"$uniqueCount"}  } } ]
			 	, function(err,result) {
		         	var i, count;

			        rowNum = rowNum + 2; 

				 	PurchaseYear.set(1, rowNum, 'Customer Count by Year');    
					PurchaseYear.set(1, rowNum + 2, 'Purchase Year');
			  		PurchaseYear.set(2, rowNum + 2, 'Customer Count');

					rowNum = rowNum + 3;


		            if(err) throw err;    
		            
		            for(i=0, count = result.length; i < count; i++) {
		            	PurchaseYear.set(1, i+rowNum, result[i]._id);
		            	PurchaseYear.set(2, i+rowNum, result[i].uniqueCustomerCount);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

		            //getRepeatTrans();
					getAverageTransValue();

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	}


	 	function getAverageTransValue() {

		 	collection.aggregate( [{ $group : { _id : "$purchase_year", avg_value: {$avg : "$billings"} } }]
			 	, function(err,result) {
		         	var i, count;

			        rowNum = rowNum + 2; 

				 	PurchaseYear.set(1, rowNum, 'Average Transaction Value');    
					PurchaseYear.set(1, rowNum + 2, 'Purchase Year');
			  		PurchaseYear.set(2, rowNum + 2, 'Average Transaction Value');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            for(i=0, count = result.length; i < count; i++) {
		            	PurchaseYear.set(1, i+rowNum, result[i]._id);
		            	PurchaseYear.set(2, i+rowNum, result[i].avg_value);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

		            processCummAvg();
					

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	}

	 		function processCummAvg() {

		 	collection.aggregate( [{ $group : { _id : "$purchase_year", avg_value: {$avg : "$billings"} } }]
			 	, function(err,result) {
		         	var i, count, cummValue;

			        rowNum = rowNum + 2; 

				 	PurchaseYear.set(1, rowNum, 'Cummulative Average Transaction Value');    
					PurchaseYear.set(1, rowNum + 2, 'Purchase Year');
			  		PurchaseYear.set(2, rowNum + 2, 'Cummulative Average Transaction Value');

					rowNum = rowNum + 3;

		            if(err) throw err;    
		            
		            cummValue = 0;
		            for(i=0, count = result.length; i < count; i++) {
		            	cummValue = cummValue + result[i].avg_value;
		            	PurchaseYear.set(1, i+rowNum, result[i]._id);
		            	PurchaseYear.set(2, i+rowNum, cummValue);
		            }

		            rowNum = rowNum + result.length;
		            //console.log(items);
		            console.log(rowNum);

		            CountryData.processCountry( collection, rowNum, CountrySheet, workbook );					

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	}


/*
	 	function processCountry() {
				 	 
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

            processCountryCC();

            // Save it 
			workbook.save(function(ok){
			   	if (!ok) 
			      workbook.cancel();
			    else
			      console.log('congratulations, your workbook created');
			});

        });
     }
        
     		
     		function processCountryCC() {

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

		            // Save it 
					workbook.save(function(ok){
					   	if (!ok) 
					      workbook.cancel();
					    else
					      console.log('congratulations, your workbook created');
					});
			 	});
	 	}
	 	*/

	 		processPurchaseYear();

		});

	});


// catch 404 and forward to error handler
app.use(function(req, res, next) {
  var err = new Error('Not Found');
  err.status = 404;
  next(err);
});

// error handler
app.use(function(err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

module.exports = app;
