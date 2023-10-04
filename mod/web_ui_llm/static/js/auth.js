import { initializeApp } from 'firebase/app';
import {
    onAuthStateChanged,
    signInWithEmailAndPassword,
    getAuth
} from 'firebase/auth';

const firebaseConfig = {
    apiKey: "AIzaSyA2pCYT_RprTqsmcVlbrURwPem-sP3dDVQ",
    authDomain: "testing-bigquery-vertexai.firebaseapp.com"
};
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);

document.addEventListener("DOMContentLoaded", () => {
    onAuthStateChanged(auth, (user) => {
        if (user) {
            document.getElementById("message").innerHTML = "Welcome, " + user.email;
        }
        else {
            document.getElementById("message").innerHTML = "No user signed in.";
        }
    });
    signIn();
});

function signIn() {
    const email = "jaronchan123@gmail.com";
    const password = "jaronchan-kpmg";
    signInWithEmailAndPassword(auth, email, password)
        .catch((error) => {
            document.getElementById("message").innerHTML = error.message;
        });
}