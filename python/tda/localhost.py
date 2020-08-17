from flask import Flask
from flask import request

app = Flask(__name__)

@app.route('/')
def hello_world():
    print(request.headers)
    print(request.data)
    return 'Hello, World!'
    
if __name__ == '__main__':
    app.run(host="localhost", port=8080, debug=True)