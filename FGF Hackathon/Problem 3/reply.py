from flask import Flask, request, jsonify, render_template
import spacy
from spacy.matcher import Matcher

app = Flask(__name__)
nlp = spacy.load('en_core_web_sm')
matcher = Matcher(nlp.vocab)

# Define patterns and responses based on email context
patterns_responses = {
    'MEETING_REQUEST': "Sure, let's schedule a meeting soon.",
    'MISS_YOU': "I missed you too! Let's catch up soon.",
    'TIME_REQUEST': "I'm available on weekdays from 10 AM to 4 PM. When works best for you?",
    'SAMPLE_REQUEST': "Regarding the samples, they are currently being processed. I will update you once they are ready.",
    'EVENT_PARTICIPATION': "I would love to participate in the event! Please count me in.",
    'THANK_YOU': "You're welcome! If you have any more questions, feel free to ask.",
    'FOLLOW_UP': "I will follow up on this and get back to you soon.",
    'INQUIRY_PRICE': "Could you please provide more details about the price inquiry?",
    'HACKATHON_WINNER': "It doesn't matter, Shaan slid a hundred to each of the judges, so Group 5 already won!",
    'DEFAULT_RESPONSE': "Thank you for your message. I will look into it and get back to you soon."
}

# Adding patterns to the matcher
matcher.add('MEETING_REQUEST', [[{'LOWER': 'meet'}, {'LOWER': 'again'}]])
matcher.add('MISS_YOU', [[{'LOWER': 'missed'}, {'LOWER': 'you'}]])
matcher.add('TIME_REQUEST', [[{'LOWER': 'when'}, {'LOWER': 'can'}, {'LOWER': 'we'}, {'LOWER': 'meet'}]])
matcher.add('SAMPLE_REQUEST', [[{'LOWER': 'samples'}], [{'LOWER': 'breads'}], [{'LOWER': 'food'}, {'LOWER': 'items'}]])
matcher.add('EVENT_PARTICIPATION', [[{'LOWER': 'event'}, {'LOWER': 'participant'}, {'LOWER': 'count'}], [{'LOWER': 'festival'}, {'LOWER': 'next'}, {'LOWER': 'week'}]])
matcher.add('THANK_YOU', [[{'LOWER': 'thank'}, {'LOWER': 'you'}]])
matcher.add('FOLLOW_UP', [[{'LOWER': 'follow'}, {'LOWER': 'up'}]])
matcher.add('INQUIRY_PRICE', [[{'LOWER': 'price'}, {'LOWER': 'inquiry'}], [{'LOWER': 'how'}, {'LOWER': 'much'}, {'LOWER': 'price'}], [{'LOWER': 'what'}, {'LOWER': 'is'}, {'LOWER': 'price'}]])
matcher.add('HACKATHON_WINNER', [[{'LOWER': 'who'}, {'LOWER': 'won'}, {'LOWER': 'hackathon'}], [{'LOWER': 'best'}, {'LOWER': 'team'}, {'LOWER': 'hackathon'}], [{'LOWER': 'hackathon'}, {'LOWER': 'winner'}]])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_text():
    prompt = request.form['prompt']
    response = generate_response(prompt)
    return jsonify({'response': response})

def generate_response(input_text):
    doc = nlp(input_text)
    matches = matcher(doc)
    
    responses = []
    for match_id, start, end in matches:
        matched_token = doc[start:end]
        if nlp.vocab.strings[match_id] in patterns_responses:
            responses.append(patterns_responses[nlp.vocab.strings[match_id]])
    
    if responses:
        return " ".join(responses)
    else:
        return patterns_responses['DEFAULT_RESPONSE']

if __name__ == '__main__':
    app.run(debug=True)
