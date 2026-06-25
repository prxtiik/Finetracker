// firebase-init.js
// FineTracker — Firebase setup
// ------------------------------------------------------------------
// Replace the values in firebaseConfig below with YOUR project's config.
// You'll get this from: Firebase Console -> Project Settings -> General
// -> "Your apps" -> Web app -> SDK setup and configuration
// ------------------------------------------------------------------

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getFirestore,
  doc,
  getDoc,
  setDoc,
  updateDoc,
  onSnapshot,
  collection,
  arrayUnion,
  serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyD7pzPD43o8wEEhBJVb5s0aUdBKRmVNeEQ",
  authDomain: "finetracker-b3bce.firebaseapp.com",
  projectId: "finetracker-b3bce",
  storageBucket: "finetracker-b3bce.firebasestorage.app",
  messagingSenderId: "705308338177",
  appId: "1:705308338177:web:4e62cc10d3438205393800"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

// Expose to app.js via window (keeps this a plain script-tag friendly setup,
// no bundler needed)
window.__fb = {
  db,
  doc,
  getDoc,
  setDoc,
  updateDoc,
  onSnapshot,
  collection,
  arrayUnion,
  serverTimestamp
};

window.dispatchEvent(new Event("firebase-ready"));
