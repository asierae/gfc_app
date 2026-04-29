import { initializeApp } from 'firebase/app';
import { getFirestore } from 'firebase/firestore';

const firebaseConfig = {
  apiKey: 'AIzaSyDV4xqyeU5lAw3ve7qqwi6eU8d3zK7qIVg',
  authDomain: 'cfg-haty.firebaseapp.com',
  projectId: 'cfg-haty',
  storageBucket: 'cfg-haty.firebasestorage.app',
  messagingSenderId: '458200679571',
  appId: '1:458200679571:web:694d75df8ee585fa723966'
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
