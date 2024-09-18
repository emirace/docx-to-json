from flask import Flask, request, jsonify
from flask_cors import CORS
from pymongo import MongoClient
import cloudinary
import cloudinary.uploader
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

app = Flask(__name__)
CORS(app)  # To handle cross-origin requests

# MongoDB setup using environment variable for connection string
MONGO_URI = os.getenv('MONGO_URI')
client = MongoClient(MONGO_URI)
db = client['mydatabase']
collection = db['mycollection']

# Cloudinary setup using environment variables
cloudinary.config(
  cloud_name=os.getenv('CLOUDINARY_CLOUD_NAME'),
  api_key=os.getenv('CLOUDINARY_API_KEY'),
  api_secret=os.getenv('CLOUDINARY_API_SECRET')
)

# Route to store JSON data in MongoDB
@app.route('/api/store-json', methods=['POST'])
def store_json():
    try:
        data = request.get_json()  # Get the JSON data from the request
        if not data:
            return jsonify({"error": "No data provided"}), 400

        # Insert the data into MongoDB
        result = collection.insert_one(data)

        return jsonify({"message": "Data stored successfully", "id": str(result.inserted_id)}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Route to upload image to Cloudinary
@app.route('/api/upload-image', methods=['POST'])
def upload_image():
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file part"}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({"error": "No selected file"}), 400

        # Upload the image to Cloudinary
        upload_result = cloudinary.uploader.upload(file)

        return jsonify({"message": "Image uploaded successfully", "url": upload_result['secure_url']}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
