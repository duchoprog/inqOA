const { MongoClient, ObjectId, GridFSBucket } = require("mongodb");
require("dotenv").config();
const readline = require("readline");

// Replace with your MongoDB Atlas connection string
const uri = process.env.MONGO_CONNECTION_STRING;
const client = new MongoClient(uri);
let mongoRecordId;

async function createMongoObject(params) {
  try {
    await client.connect();
    console.log("Connected to MongoDB Atlas");

    const database = client.db("inquiry");
    const collection = database.collection("inquiryTempStorage");

    // Create a document
    const insertResult = await collection.insertOne({
      files: [],
      excelBase: [],
      images: [],
      imageVault: [],
      output: [],
      uploads: [],
    });
    //console.log("Inserted document =>", insertResult.insertedId);
    return insertResult.insertedId;
  } catch (error) {
    console.log(error);
  }
}
async function findOneMongo(id) {
  console.log("findonemongo id", id);

  try {
    await client.connect();
    const database = client.db("inquiry");
    const collection = database.collection("inquiryTempStorage");
    if (id) {
      const findResult = await collection.findOne({
        _id: id,
      });
      console.log("Found document by _id =>", findResult);
      return findResult;
    } else {
      console.log("id not found");
      throw new Error("id not found");
    }
  } catch (error) {
    console.log(error);
  } finally {
    await client.close();
    console.log("Connection to MongoDB Atlas closed");
  }
}
///////
async function findAllMongo() {
  try {
    await client.connect();
    const database = client.db("inquiry");
    const collection = database.collection("inquiryTempStorage");
    await collection
      .find({})
      .toArray()
      .then((ans) => {
        console.log(ans);
      });
    //console.log("Found all =>", findResult);
    return findResult;
  } catch (error) {
    console.log(error);
  } finally {
    await client.close();
    console.log("Connection to MongoDB Atlas closed");
  }
}

//////
// Function to update MongoDB object
async function addToMongoObject(id, key, fileContent) {
  try {
    await client.connect();
    const database = client.db("inquiry");
    const collection = database.collection("inquiryTempStorage");

    const result = await collection.updateOne(
      { _id: id },
      { $push: { [key]: fileContent } }
    );

    console.log("Updated document =>", result);
  } catch (error) {
    console.error("Error updating MongoDB object:", error);
  } finally {
    await client.close();
  }
}
//////////

async function storeImageFromMemory(imageBuffer, fileName) {
  await client.connect();

  const database = client.db("inquiry");
  const bucket = new GridFSBucket(db, { bucketName: "images" });

  const uploadStream = bucket.openUploadStream(fileName);
  uploadStream.end(imageBuffer);

  uploadStream.on("error", function (error) {
    console.error("Error uploading file:", error);
  });

  uploadStream.on("finish", function () {
    console.log("File uploaded successfully");
    client.close();
  });
}

// Example usage
/* const imageBuffer = Buffer.from('your_image_data', 'base64'); // Replace with your actual image buffer
storeImageFromMemory(imageBuffer, 'image.jpg'); */

/////////////

async function main() {
  //const client = new MongoClient(uri);

  try {
    await createMongoObject();
    console.log("Inserted document =>", insertResult.insertedId);
    await pauseExecution();
    // Read documents
    const findResult = await collection.find({}).toArray();
    console.log("Found documents =>", findResult);
    await pauseExecution();
    // Update a document
    const updateResult = await collection.updateOne(
      { _id: ObjectId(insertResult.insertedId) },
      { $set: { age: 31 } }
    );
    console.log("Updated document =>", updateResult.modifiedCount);
    await pauseExecution();
    // Delete a document
    const deleteResult = await collection.deleteOne({
      _id: ObjectId(insertResult.insertedId),
    });
    console.log("Deleted document =>", deleteResult.deletedCount);
    await pauseExecution();
  } finally {
    await client.close();
    console.log("Connection to MongoDB Atlas closed");
    await pauseExecution();
  }
}

function pauseExecution() {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  return new Promise((resolve) => {
    rl.question("Press any key to continue...", () => {
      rl.close();
      resolve();
    });
  });
}
module.exports = {
  createMongoObject,
  findOneMongo,
  addToMongoObject,
  findAllMongo,
};
