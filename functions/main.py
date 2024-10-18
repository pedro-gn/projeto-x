import io
import pathlib
import re
from firebase_functions import https_fn, options
from firebase_admin import initialize_app, storage
from docx import Document
from flask import jsonify, make_response
import pandas as pd
from util import procura_marcacoes, matches_dict, replace_doc
import json
PADRAO = r"\{\{(.+?)\}\}"

# Initialize Firebase app
initialize_app()


@https_fn.on_request()
def get_matches(req: https_fn.Request) -> https_fn.Response:
    # Handle preflight request
    if req.method == 'OPTIONS':
        response = make_response('', 204)
        response.headers['Access-Control-Allow-Origin'] = '*'
        response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
        response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return response

    try:
        # Parse JSON request data
        request_json = req.get_json()
        if not request_json:
            response = make_response(jsonify({"error": "Invalid JSON request"}), 400)
            response.headers['Access-Control-Allow-Origin'] = '*'
            return response

        bucket_name = request_json.get("bucket_name")
        template_file_path = request_json.get("template_file")
        spreadsheet_file_path = request_json.get("spreadsheet_file")

        if not bucket_name or not template_file_path or not spreadsheet_file_path:
            response = make_response(jsonify({"error": "Missing required parameters"}), 400)
            response.headers['Access-Control-Allow-Origin'] = '*'
            return response

        # Get the bucket
        bucket = storage.bucket(bucket_name)

        # Get the template file from the bucket
        template_blob = bucket.blob(template_file_path)
        template_file_content = template_blob.download_as_string()

        # Load the template file into a Document object
        template_document = Document(io.BytesIO(template_file_content))

        # Get the spreadsheet file from the bucket
        spreadsheet_blob = bucket.blob(spreadsheet_file_path)
        spreadsheet_file_content = spreadsheet_blob.download_as_string()

        # Load the spreadsheet file into a pandas DataFrame
        spreadsheet_df = pd.read_excel(io.BytesIO(spreadsheet_file_content))

        # Process the files 
        matches = procura_marcacoes(template_document)
        print(matches)
        final_dict = matches_dict(template_document, matches, spreadsheet_df)

        response = make_response(jsonify(json.dumps(final_dict)), 200)
        response.headers['Access-Control-Allow-Origin'] = '*'
        return response

    except Exception as e:
        print("Error processing files:", e)
        response = make_response(jsonify({"error": "Internal Server Error"}), 500)
        response.headers['Access-Control-Allow-Origin'] = '*'
        return response







@https_fn.on_request()
def process_files(req: https_fn.Request) -> https_fn.Response:
    # Handle preflight request
    if req.method == 'OPTIONS':
        response = make_response('', 204)
        response.headers['Access-Control-Allow-Origin'] = '*'
        response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
        response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return response

    try:
        # Parse JSON request data
        request_json = req.get_json()
        if not request_json:
            response = make_response(jsonify({"error": "Invalid JSON request"}), 400)
            response.headers['Access-Control-Allow-Origin'] = '*'
            return response

        bucket_name = request_json.get("bucket_name")
        template_file_path = request_json.get("template_file")
        spreadsheet_file_path = request_json.get("spreadsheet_file")
        matches =  request_json.get("matches")


        if not bucket_name or not template_file_path or not spreadsheet_file_path:
            response = make_response(jsonify({"error": "Missing required parameters"}), 400)
            response.headers['Access-Control-Allow-Origin'] = '*'
            return response

        # Get the bucket
        bucket = storage.bucket(bucket_name)

        # Get the template file from the bucket
        template_blob = bucket.blob(template_file_path)
        template_file_content = template_blob.download_as_string()

        # Load the template file into a Document object
        template_document = Document(io.BytesIO(template_file_content))

        # Get the spreadsheet file from the bucket
        spreadsheet_blob = bucket.blob(spreadsheet_file_path)
        spreadsheet_file_content = spreadsheet_blob.download_as_string()

        # Load the spreadsheet file into a pandas DataFrame
        spreadsheet_df = pd.read_excel(io.BytesIO(spreadsheet_file_content))

        doc_path = pathlib.Path(template_file_path)
        
        replace_doc(template_document, matches, spreadsheet_df, doc_path, bucket)

        response = make_response(jsonify({"matches": "Concluido"}), 200)
        response.headers['Access-Control-Allow-Origin'] = '*'
        return response

    except Exception as e:
        print("Error processing files:", e)
        response = make_response(jsonify({"error": "Internal Server Error"}), 500)
        response.headers['Access-Control-Allow-Origin'] = '*'
        return response