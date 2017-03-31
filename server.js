// mongo-express-node-example/server.js

// An Express.js Web app that performs CRUD operations on a Mongo database

// TODO: 2017/03/29 : Eliminate the abundance of redundancy in this code. Code should be DRY.

// "dependencies": {
    // "body-parser": "~1.12.4",
    // "cookie-parser": "~1.3.5",
    // "debug": "~2.2.0",
    // "express": "~4.12.4",
    // "jade": "~1.9.2",
    // "morgan": "~1.5.3",
    // "serve-favicon": "~2.2.1",
    // "mongodb": "^1.4.4",
    // "monk": "^1.0.1"
// }

//import xlsx from 'node-xlsx';
// Or var xlsx = require('node-xlsx').default; 
var xlsx = require('node-xlsx').default; 

// From https://blog.xervo.io/nodejs-and-express-create-rest-api :

var express = require('express');
var path = require('path');
var favicon = require('serve-favicon');
// var logger = require('morgan');
// var cookieParser = require('cookie-parser');
// var bodyParser = require('body-parser');

// var routes = require('./routes/index');
// var users = require('./routes/users');

var app = express();

// BEGIN: Stolen (in a good way) from http://cwbuecheler.com/web/tutorials/2013/node-express-mongo/

// view engine setup
// app.set('views', path.join(__dirname, 'views'));
// app.set('view engine', 'jade');

app.use(favicon(__dirname + '/public/favicon.ico'));

// app.use(logger('dev'));
// app.use(bodyParser.json());
// app.use(bodyParser.urlencoded({ extended: false }));
// app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

// app.use('/', routes);
// app.use('/users', users);

// END: Stolen...

app.get('/xlsx_from_ram', function(req, res) {
	const data = [[1, 2, 3], [true, false, null, 'sheetjs'], ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'], ['baz', null, 'qux']];
	var buffer = xlsx.build([{name: "mySheetName", data: data}]); // Returns a buffer
	var filename = 'test.xlsx';

	// res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); // See https://stackoverflow.com/questions/4212861/what-is-a-correct-mime-type-for-docx-pptx-etc

	res.set({ // See https://stackoverflow.com/questions/14897975/nodejs-express-send-file
		"Content-Disposition": 'attachment; filename="' + filename + '"',
		"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" //,
		// "Content-Length": data.Body.length
	});

	res.send(buffer);
});

app.get('/find_all', function(req, res) {
	// From https://www.npmjs.com/package/mongodb

	var MongoClient = require('mongodb').MongoClient;
	var assert = require('assert');
	 
	// Connection URL 
	var url = 'mongodb://localhost:27017/express-xlsx-response-test';

	var findDocuments = function(db, callback) {
		// Get the usercollection collection 
		var collection = db.collection('usercollection');

		collection.find({}).toArray(function(err, docs) {
			assert.equal(err, null);
			assert.equal(3, docs.length);
			console.log("Found the following records:");
			console.dir(docs);
			console.log("Record number 0:", docs[0]);
			console.log("Record number 0 username:", docs[0].username);
			console.log("Record number 0 email:", docs[0].email);
			
			docs.forEach(function(item, index) {
				console.log("Record number", index , "username:", item.username);
				console.log("Record number", index , "email:", item.email);
			});
			
			callback(docs);
		});
	}

	// Use connect method to connect to the Server 
	MongoClient.connect(url, function(err, db) {
		assert.equal(null, err);
		console.log("Connected to MongoDB");

		findDocuments(db, function(docs) {
			db.close();
			console.log("Disconnected from MongoDB");

			// var data = [[1, 2, 3], [true, false, null, 'sheetjs']];
			var data = [['Header 1', 'Header 2', 'Header 3']];

			docs.forEach(function(item, index) {
				data.push([index, item.username, item.email]);
			});

			var buffer = xlsx.build([{name: "mySheetName", data: data}]); // Returns a buffer
			var filename = 'mongodb.xlsx';

			res.set({ // See https://stackoverflow.com/questions/14897975/nodejs-express-send-file
				"Content-Disposition": 'attachment; filename="'+filename+'"',
				"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" //,
				// "Content-Length": data.Body.length
			});

			console.log("Sending the HTTP response...");
			res.send(buffer);
			console.log("Sent the HTTP response.");
		});
	});
});

app.get('/find_one', function(req, res) {
	// From https://www.npmjs.com/package/mongodb

	var MongoClient = require('mongodb').MongoClient;
	var assert = require('assert');
	 
	// Connection URL 
	var url = 'mongodb://localhost:27017/express-xlsx-response-test';

	var findDocuments = function(db, callback) {
		// Get the usercollection collection 
		var collection = db.collection('usercollection');

		collection.find({ username: 'testuser4' }).toArray(function(err, docs) {
			assert.equal(err, null);
			assert.equal(1, docs.length);
			console.log("Found the following records:");
			console.dir(docs);
			
			docs.forEach(function(item, index) {
				console.log("Record number", index , "username:", item.username);
				console.log("Record number", index , "email:", item.email);
			});
			
			callback(docs);
		});
	}

	// Use connect method to connect to the Server 
	MongoClient.connect(url, function(err, db) {
		assert.equal(null, err);
		console.log("Connected to MongoDB");

		findDocuments(db, function(docs) {
			db.close();
			console.log("Disconnected from MongoDB");

			// var data = [[1, 2, 3], [true, false, null, 'sheetjs']];
			var data = [['Header 1', 'Header 2', 'Header 3']];

			docs.forEach(function(item, index) {
				data.push([index, item.username, item.email]);
			});

			var buffer = xlsx.build([{name: "mySheetName", data: data}]); // Returns a buffer
			var filename = 'mongodb.xlsx';

			res.set({ // See https://stackoverflow.com/questions/14897975/nodejs-express-send-file
				"Content-Disposition": 'attachment; filename="'+filename+'"',
				"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" //,
				// "Content-Length": data.Body.length
			});

			console.log("Sending the HTTP response...");
			res.send(buffer);
			console.log("Sent the HTTP response.");
		});
	});
});

app.get('/insert', function(req, res) {
	// From https://www.npmjs.com/package/mongodb

	var MongoClient = require('mongodb').MongoClient;
	var assert = require('assert');
	 
	// Connection URL 
	var url = 'mongodb://localhost:27017/express-xlsx-response-test';

	var insertDocument = function(db, callback) {
		// Get the documents collection 
		var collection = db.collection('usercollection');
		collection.insert(
			{ "username" : "testuser4", "email" : "testuser4@testdomain.com" },
			function(err, result) {
				console.log('Result of insert:', result);
				assert.equal(err, null);
				assert.equal(1, result.result.n);
				assert.equal(1, result.ops.length);
				console.log("Insertion successful.");
				// callback(result);
				callback(); // (err);
			});
	}

	// Use connect method to connect to the Server 
	MongoClient.connect(url, function(err, db) {
		assert.equal(null, err);
		console.log("Connected to MongoDB");

		insertDocument(db, function() /* (err2) */ {
			db.close();
			console.log("Disconnected from MongoDB");

			// if (err2) { Set the status code to 500 Internal Server Error }
			
			console.log("Sending the HTTP response...");
			res.send(); // HTTP status code: 200 OK
			console.log("Sent the HTTP response.");
		});
	});
});

app.get('/update', function(req, res) {
	// From https://www.npmjs.com/package/mongodb

	var MongoClient = require('mongodb').MongoClient;
	var assert = require('assert');
	 
	// Connection URL 
	var url = 'mongodb://localhost:27017/express-xlsx-response-test';

	var updateDocument = function(db, callback) {
		// Get the documents collection 
		var collection = db.collection('usercollection');
		// Update document where a is 2, set b equal to 1 
		collection.updateOne(
			{ username : 'testuser2' },
			{ $set: { email : 'funky@chicken.org' } },
			function(err, result) {
				assert.equal(err, null);
				assert.equal(1, result.result.n);
				console.log("Updated the document with the field 'username' equal to 'testuser2'");
				// callback(result);
				callback(); // (err);
			});  
	}

	// Use connect method to connect to the Server 
	MongoClient.connect(url, function(err, db) {
		assert.equal(null, err);
		console.log("Connected to MongoDB");

		updateDocument(db, function() /* (err2) */ {
			db.close();
			console.log("Disconnected from MongoDB");

			// if (err2) { Set the status code to 500 Internal Server Error }
			
			console.log("Sending the HTTP response...");
			res.send(); // HTTP status code: 200 OK
			console.log("Sent the HTTP response.");
		});
	});
});

app.get('/delete', function(req, res) {
	// From https://www.npmjs.com/package/mongodb

	var MongoClient = require('mongodb').MongoClient;
	var assert = require('assert');
	 
	// Connection URL 
	var url = 'mongodb://localhost:27017/express-xlsx-response-test';

	var deleteDocument = function(db, callback) {
		var collection = db.collection('usercollection');
		collection.deleteOne(
			{ username : 'testuser3' },
			function(err, result) {
				assert.equal(err, null);
				assert.equal(1, result.result.n);
				console.log("Removed the document with the field 'username' equal to 'testuser3'");
				// callback(result);
				callback(); // (err);
			});
	}

	// Use connect method to connect to the Server 
	MongoClient.connect(url, function(err, db) {
		assert.equal(null, err);
		console.log("Connected to MongoDB");

		deleteDocument(db, function() /* (err2) */ {
			db.close();
			console.log("Disconnected from MongoDB");

			// if (err2) { Set the status code to 500 Internal Server Error }
			
			console.log("Sending the HTTP response...");
			res.send(); // HTTP status code: 200 OK
			console.log("Sent the HTTP response.");
		});
	});
});

app.listen(3000);

// ---

// To set up the MongoDB database, launch mongod.exe (the database "daemon") and mongo.exe (the MongoDB shell), and type the following into the shell:

// use express-xlsx-response-test
// newstuff = [{ "username" : "testuser1", "email" : "testuser1@testdomain.com" }, { "username" : "testuser2", "email" : "testuser2@testdomain.com" }, { "username" : "testuser3", "email" : "testuser3@testdomain.com" }]
// db.usercollection.insert(newstuff);
// db.usercollection.find().pretty()
// db.stats()
// show dbs
// show collections

// The End.
