// Firebase Config

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";

import {
  getAuth
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";

import {
  getFirestore
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyCLaEjD4xQImoHJVVjpND5hKFgqHswLcmI",
  authDomain: "hrms-portal-aab0b.firebaseapp.com",
  projectId: "hrms-portal-aab0b",
  storageBucket: "hrms-portal-aab0b.firebasestorage.app",
  messagingSenderId: "229143989837",
  appId: "1:229143989837:web:d6a122c31f41c077225604",
  measurementId: "G-76103ZL5YH"
};

const app = initializeApp(firebaseConfig);

export const auth = getAuth(app);

export const db = getFirestore(app);
