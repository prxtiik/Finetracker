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
  apiKey: "REPLACE_ME",
  authDomain: "REPLACE_ME.firebaseapp.com",
  projectId: "REPLACE_ME",
  storageBucket: "REPLACE_ME.appspot.com",
  messagingSenderId: "REPLACE_ME",
  appId: "REPLACE_ME"
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
