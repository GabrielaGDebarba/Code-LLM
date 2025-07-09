import sys
sys.path.insert(0, "C:/Code-LLM" )


from flask import Flask, request, jsonify
from config import *
from llm_calls import *

app = Flask(__name__)

# isto foi uma mudancass


@app.route('/get_parameters', methods=['POST'])
def llm_call():
    data = request.get_json()
    input_string = data.get('input', '')

    answer = make_floorplan(input_string) 
    print(answer)

    return answer

if __name__ == '__main__':
    app.run(debug=True)

