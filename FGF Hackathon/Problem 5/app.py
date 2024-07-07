import os
import nltk
import datetime
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import SVC
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler
from nltk.stem import WordNetLemmatizer
import string
from flask import Flask, request, render_template

app = Flask(__name__)

# Initialize NLTK stopwords and tokenizer
nltk.download('stopwords')
nltk.download('punkt')
nltk.download('wordnet')

stop_words = set(stopwords.words('english'))
lemmatizer = WordNetLemmatizer()

# Example dataset (email texts and corresponding categories)
emails = [
    ("Urgent", "Urgent"),
    ("Informational Project Update", "Informational"),
    ("Action Required: Review Proposal", "Actionable"),
    ("Please forward this email to the technical team", "Delegate"),
    ("Requesting technical assistance", "Delegate"),
    ("Meeting rescheduled to next week", "Informational"),
    ("Immediate action needed on the project", "Urgent"),
    ("Follow up on the proposal review", "Actionable"),
    ("Client query regarding the new feature", "Delegate"),
    ("Weekly team meeting minutes", "Informational"),
    ("Reminder: Submit your timesheets", "Actionable"),
    ("Customer feedback on recent purchase", "Delegate"),
    ("Budget review meeting scheduled", "Informational"),
    ("Final notice: Payment overdue", "Urgent"),
    ("Project kickoff meeting agenda", "Informational"),
    ("Update required: Project status", "Actionable"),
    ("Escalate: Critical issue reported", "Urgent"),
    ("Technical documentation request", "Delegate"),
    ("New policy announcement", "Informational"),
    ("Please review the attached document", "Actionable"),
    ("Request for additional resources", "Delegate"),
    ("Annual performance reviews", "Informational"),
    ("Urgent: Security breach detected", "Urgent"),
    ("Invitation to company webinar", "Informational"),
    ("Clarification needed on the project", "Actionable"),
    ("Vendor request for proposal", "Delegate"),
    ("Monthly financial report", "Informational"),
    ("Critical update: System downtime", "Urgent"),
    ("Please complete the survey", "Actionable"),
    ("Question regarding product features", "Delegate"),
    ("Training session next week", "Informational"),
    ("Immediate attention: Compliance issue", "Urgent"),
    ("Approval needed for the budget", "Actionable"),
    ("Follow-up on previous email", "Delegate"),
    ("Notice: Office relocation", "Informational"),
    ("Urgent: Server outage", "Urgent"),
    ("Request for meeting schedule", "Actionable"),
    ("Information on new software release", "Informational"),
    ("Feedback requested on the new design", "Delegate"),
    ("Notice of holiday schedule", "Informational"),
    ("Update: New hire onboarding process", "Informational"),
    ("Please approve the attached invoice", "Actionable"),
    ("Question about the contract terms", "Delegate"),
    ("Reminder: Upcoming event", "Informational"),
    ("Urgent: Customer complaint", "Urgent"),
    ("Schedule for the next sprint planning", "Informational"),
    ("Action required: Update your contact information", "Actionable"),
    ("Forward this email to the sales team", "Delegate"),
    ("Information: Quarterly business review", "Informational"),
    ("Attention needed: Licensing issue", "Urgent"),
    ("Sign and return the attached agreement", "Actionable"),
    ("Inquiry about service availability", "Delegate"),
    ("Announcement: New office location", "Informational"),
    ("Alert: High priority bug found", "Urgent"),
    ("Request for proposal submission", "Actionable"),
    ("Please review the client feedback", "Delegate"),
    ("Update: Upcoming changes to the policy", "Informational"),
    ("Urgent: Immediate response required", "Urgent"),
    ("Fill out the attached form", "Actionable"),
    ("Forward to HR department", "Delegate"),
    ("Notice: Change in office hours", "Informational"),
]

# Additional keywords for each category
keywords = {
    "Urgent": ["urgent", "deadline", "immediate", "critical", "escalate", "attention", "asap", "priority", "emergency"],
    "Informational": ["information", "update", "meeting", "announcement", "briefing", "newsletter", "summary", "report"],
    "Actionable": ["action", "required", "follow up", "response", "complete", "deadline", "reply", "schedule", "confirm"],
    "Delegate": ["forward", "query", "assist", "request", "delegate", "assign", "redirect", "send", "transfer"]
}

# Preprocessing function
def preprocess(text):
    text = text.translate(str.maketrans('', '', string.punctuation))
    tokens = nltk.word_tokenize(text.lower())
    tokens = [lemmatizer.lemmatize(token) for token in tokens]
    tokens = [token for token in tokens if token not in stop_words]
    return ' '.join(tokens)

# Preprocess the dataset using keywords
preprocessed_emails = []
for text, label in emails:
    preprocessed_text = preprocess(text)
    for category, category_keywords in keywords.items():
        if any(keyword in preprocessed_text for keyword in category_keywords):
            preprocessed_emails.append((preprocessed_text, label))
            break

# Feature extraction (TF-IDF)
tfidf_vectorizer = TfidfVectorizer()
X = tfidf_vectorizer.fit_transform([text for text, label in preprocessed_emails])
y = [label for text, label in preprocessed_emails]

# Train SVM classifier
svm_classifier = make_pipeline(StandardScaler(with_mean=False), SVC(kernel='linear'))
svm_classifier.fit(X, y)

# Function to categorize email
def categorize_email(email_text):
    preprocessed_text = preprocess(email_text)
    text_vectorized = tfidf_vectorizer.transform([preprocessed_text])
    predicted_category = svm_classifier.predict(text_vectorized)[0]
    return predicted_category

# Flask route for home page
@app.route('/')
def home():
    return render_template('index.html')

# Flask route for handling form submission
@app.route('/submit', methods=['POST'])
def submit():
    email_content = request.form['email_content']
    category = categorize_email(email_content)

    # Save email content to file under appropriate category folder with timestamp
    folder_path = os.path.join('categorized_emails', category)
    os.makedirs(folder_path, exist_ok=True)
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    file_path = os.path.join(folder_path, f'email_{timestamp}.txt')
    with open(file_path, 'w') as file:
        file.write(email_content + "\n\n")

    return render_template('result.html', category=category)

if __name__ == '__main__':
    app.run(debug=True)
