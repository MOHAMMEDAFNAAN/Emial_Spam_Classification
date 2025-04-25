import pythoncom
import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
from win32com.client import Dispatch

# Function to speak the result
def speak(text):
    pythoncom.CoInitialize()  # Initialize COM
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

# Load model and vectorizer
model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))

# Main Streamlit app
def main():
    st.title("📧 Email Spam Classification Application")
    st.write("Built with Streamlit & Python")

    activities = ["Classification", "About"]
    choice = st.sidebar.selectbox("Select Activity", activities)

    if choice == "Classification":
        st.subheader("Enter your email/message text below:")
        msg = st.text_input("Email Content")

        if st.button("Classify"):
            if msg.strip() == "":
                st.warning("⚠️ Please enter some text to classify.")
            else:
                data = [msg]
                vec = cv.transform(data).toarray()
                result = model.predict(vec)

                if result[0] == 0:
                    st.success("✅ This is Not A Spam Email")
                    speak("This is not a spam email")
                else:
                    st.error("🚫 This is A Spam Email")
                    speak("This is a spam email")

    elif choice == "About":
        st.subheader("About this app")
        st.markdown("""
        This application uses a machine learning model to classify whether a given email/message text is spam or not.
        - **Model**: Trained on a dataset of email/text messages.
        - **Vectorizer**: CountVectorizer from sklearn.
        - **Voice Feedback**: Uses Windows SAPI for speech output.
        
        Made with ❤️ using Streamlit and Python.
        """)

# Run the app
if __name__ == "__main__":
    main()
