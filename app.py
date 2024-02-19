from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
# Import my functions here
from data_processing import data_excel_json, json_to_df,convert_numpy
from agent import job_schedule


app = Flask(__name__)
CORS(app)  # Handle CORS if interacting with a frontend running on a different server/port

@app.route('/upload', methods=['POST'])
def upload_and_schedule():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file and file.filename.endswith('.xlsx'):
        # Process the Excel file
        data_dict = data_excel_json(file)
        final_schedule= job_schedule(json_to_df(data_dict), population_size=60, crossover_rate=0.8, mutation_rate=0.2, mutation_selection_rate=0.2, num_iteration=2000)
    
        # Convert NumPy int64 types to Python native int types for JSON serialization
        final_schedule_converted = convert_numpy(final_schedule)

        # Return the processed and converted schedule
        return jsonify(final_schedule_converted)

    else:
        return jsonify({'error': 'Invalid file format'}), 400

if __name__ == '__main__':
    app.run(debug=True)
