import sys
sys.path.insert(0, "C:/Code-LLM" )

from flask import Flask, request, jsonify
from config import client
from llm_calls import *

app = Flask(__name__)

@app.route('/get_parameters', methods=['POST'])
def llm_call():
    data = request.get_json()
    input_string = data.get("input", "")
    
    # Remove both outer and inner single quotes
    cleanup_input = input_string.strip("'").replace("'", "")
    answer = make_floorplan(cleanup_input)
    
    return jsonify({'response': answer})

if __name__ == '__main__':
    app.run(debug=True)

