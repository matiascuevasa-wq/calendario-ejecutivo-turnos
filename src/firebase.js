import { initializeApp } from "firebase/app";
import { getFirestore } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyAVFPItuXFeQF8u3HKgFlhrviHFCtauNns",
  authDomain: "calendario-turnos-oems.firebaseapp.com",
  projectId: "calendario-turnos-oems",
  storageBucket: "calendario-turnos-oems.firebasestorage.app",
  messagingSenderId: "600455008831",
  appId: "1:600455008831:web:dc8363a5ffd9a990ef1e8a",
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);

