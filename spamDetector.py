import pythoncom
import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
import numpy as np
from win32com.client import Dispatch

def speak(text):
    pythoncom.CoInitialize()
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(text)

# Load model and vectorizer
model = pickle.load(open('spam.pkl', 'rb'))
cv = pickle.load(open('vectorizer.pkl', 'rb'))

def main():
    st.title("📧 Email Spam Classification App")
    st.write("Built with Streamlit & Python")

    activities = ["Classification", "About"]
    choice = st.sidebar.selectbox("Select Activity", activities)

    if choice == "Classification":
        st.subheader("Classify an Email")
        msg = st.text_input("Enter your email text here:")

        if st.button("Process"):
            data = [msg]
            vec = cv.transform(data).toarray()

            # Show vector and vocab for debugging
            st.write("🔍 Vectorized Input:", vec)
            st.write("📘 Vocabulary Snapshot:", list(cv.vocabulary_.keys())[:50])

            # Predict class and confidence
            result = model.predict(vec)
            probas = model.predict_proba(vec)
            st.write(f"📊 Confidence → Not Spam: {probas[0][0]:.2f}, Spam: {probas[0][1]:.2f}")

            if result[0] == 0:
                st.success("✅ This is Not a Spam Email")
                speak("This is Not a Spam Email")
            else:
                st.error("⚠️ This is a Spam Email")
                speak("This is a Spam Email")

    elif choice == "About":
        st.subheader("About this App")
        st.write("This is a simple machine learning app that classifies emails as spam or not spam using a trained model.")

main()
