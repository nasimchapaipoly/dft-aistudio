import React, { useState, useEffect, useMemo, FormEvent } from 'react';
import * as XLSX from 'xlsx';
import { 
  Clock, 
  Leaf, 
  Menu, 
  X, 
  FlaskConical, 
  Users, 
  BookOpen, 
  Briefcase, 
  MapPin, 
  Phone, 
  Mail, 
  ChevronRight,
  Facebook,
  Youtube,
  Linkedin,
  Pin,
  Download,
  Printer,
  Search,
  ChevronDown,
  ChevronUp,
  Globe,
  FileText,
  Target,
  LogOut,
  LogIn,
  Plus,
  Upload,
  Settings,
  Building2,
  Trash2,
  User as UserIcon
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell } from 'recharts';
import { db, auth, googleProvider, storage, OperationType, handleFirestoreError } from './lib/firebase';
import { 
  collection, 
  onSnapshot, 
  query, 
  orderBy, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  doc,
  serverTimestamp,
  Timestamp,
  getDocs,
  where
} from 'firebase/firestore';
import { 
  ref, 
  uploadBytes, 
  getDownloadURL 
} from 'firebase/storage';
import { 
  signInWithPopup, 
  signOut, 
  onAuthStateChanged, 
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  sendEmailVerification,
  updateProfile
} from 'firebase/auth';
import type { User } from 'firebase/auth';

type Language = 'en' | 'bn';

interface Teacher {
  name_en: string;
  name_bn: string;
  position_en: string;
  position_bn: string;
  email: string;
  mobile: string;
  image: string;
}

interface Notice {
  date: string;
  text_en: string;
  text_bn: string;
  file?: string;
}

interface Lab {
  name_en: string;
  name_bn: string;
  desc_en: string;
  desc_bn: string;
  images: string[];
}

interface Student {
  name: string;
  roll: string;
  registration: string;
  semester: string;
  shift: string;
  group: string;
  session: string;
  mobile: string;
  image: string;
}

interface StudentProfile {
  userId: string;
  basic: {
    name: string;
    roll: string;
    registration: string;
    email: string;
    mobile: string;
    semester?: string;
    shift?: string;
    group?: string;
    session?: string;
    misId?: string;
    dob?: string;
    birthReg?: string;
    nid?: string;
    image?: string;
  };
  education: {
    institution: string;
    exam: string;
    subject: string;
    board: string;
    result: string;
    year: string;
  }[];
  parents: {
    father: {
      name: string;
      occupation: string;
      mobile: string;
      nid: string;
      dob: string;
    };
    mother: {
      name: string;
      occupation: string;
      mobile: string;
      nid: string;
      dob: string;
    };
    guardian?: {
      name: string;
      relation: string;
      mobile: string;
      occupation: string;
    };
  };
  address: {
    current: string;
    permanent: string;
  };
  professional: {
    objective: string;
    experience: string;
    certificates: string;
    languages: string;
  };
  financial: {
    bankName: string;
    accName: string;
    accNo: string;
    accNid: string;
    bkash: string;
    rocket: string;
  };
}

const translations = {
  en: {
    dept: "Food Technology",
    inst: "Chapainawabganj Polytechnic",
    home: "Home",
    about: "About Us",
    teachers: "Teachers",
    staff: "Office & Lab Staff",
    syllabus: "Syllabus",
    labs: "Labs",
    students: "Students",
    studentPortal: "Student Portal",
    notices: "Notices",
    heroTitle: "Department of Food Technology",
    heroSub: "Learn • Innovate • Lead",
    meetTeachers: "Meet Our Teachers",
    explore: "Explore More",
    latestNews: "Latest News",
    tickerMsg: "Admission is open for the 2026 academic session. • Upcoming Food Festival next week! • Semester finals routine published.",
    noticeBoard: "Notice Board",
    allNotices: "All Notices",
    noticeSearch: "Search notices...",
    whyChoose: "Why Choose Us",
    modernLabs: "Modern Labs",
    expertTeachers: "Expert Teachers",
    careerSupport: "Career Support",
    aboutTitle: "About the Department",
    aboutDesc: "The Department of Food Technology at Chapainawabganj Polytechnic Institute is dedicated to providing world-class technical education. We focus on food processing, preservation, quality control, and modern food engineering to prepare students for the global industry.",
    historyTitle: "🏫 Our History",
    historyDesc: "The Department of Food Technology at Chapainawabganj Polytechnic Institute was established with the primary goal of producing highly skilled and competent diploma engineers for the rapidly growing food industry. As the demand for safe, processed, and quality food continues to increase both nationally and globally, the need for technically trained professionals in this sector has become essential. Keeping this demand in mind, the department was founded to equip students with the necessary technical knowledge, practical skills, and industry-oriented training.\n\nChapainawabganj is widely known for its rich agricultural resources, especially mangoes and other fruits and crops. This geographical advantage created a strong potential for the development of food processing industries in the region. To utilize these local resources efficiently and to add value through modern processing techniques, the Food Technology Department was introduced as a significant step toward regional and national development.\n\nIn the beginning, the department started its journey with limited infrastructure and facilities. However, through continuous improvement, support from educational authorities, and the dedication of experienced teachers, it has gradually developed into a well-established and modern department. Today, it is equipped with updated laboratories, improved equipment, and a practical-based teaching approach that helps students gain real-life industrial experience.\n\nOver the years, the department has successfully produced many skilled graduates who are now working in various reputed food industries, both in Bangladesh and abroad. Their success reflects the quality of education and training provided by the department. As a result, the Department of Food Technology has earned a strong reputation and is now considered one of the most prestigious and promising departments in the region.\n\nThe continuous progress, commitment to excellence, and focus on practical learning have made this department a key contributor to the development of the food sector. It not only provides education but also plays an important role in building a skilled workforce for the future.",
    missionTitle: "🎯 Our Mission",
    missionDesc: "Our mission is to provide comprehensive and modern technical education that equips students with the necessary skills in food processing, preservation, and quality control, enabling them to meet industry standards effectively.\n\nTo achieve this goal, the department places strong emphasis on both theoretical knowledge and practical, hands-on training. Through modern laboratories, advanced equipment, and industry-oriented learning approaches, students gain in-depth understanding of various stages of food production and processing.\n\nIn addition, students are trained in food safety, hygiene practices, and international standards such as ISO and HACCP, so that they can perform efficiently in real-world industrial environments.\n\nAnother important objective of the department is to develop students’ professional skills, sense of responsibility, and ethical values, preparing them to become competent and successful food technologists.\n\nOverall, the mission of the department is not only to provide knowledge but also to shape students into skilled professionals who can meet the demands of the modern food industry and contribute effectively to the sector.",
    visionTitle: "🌍 Our Vision",
    visionDesc: "Our vision is to become a center of excellence in Food Technology education, contributing to innovation, quality, and sustainability in the global food sector.\n\nWe aim to establish the department as a leading and recognized institution that provides high-quality education and training aligned with international standards. By integrating modern technologies, research, and practical learning, we strive to prepare students who can compete and succeed in both national and global food industries.\n\nThe department is committed to promoting innovation in food processing, preservation, and product development. We encourage students to think creatively and develop new ideas that can improve food quality, safety, and efficiency.\n\nSustainability is also a key focus of our vision. We aim to develop environmentally friendly practices in food production and processing, ensuring responsible use of resources and reducing food waste.\n\nThrough continuous improvement, research, and collaboration with industries, our vision is to produce skilled food technologists who will contribute to the advancement of the global food sector and play a vital role in ensuring food security and safety for the future.",
    syllabusDesc: "Download the BTEB Probidhan 2022 Syllabus for Food Technology.",
    academicInfo: "Academic Information",
    facultyMsg: "Meet our exceptional faculty members leading the way in food innovation.",
    staffMsg: "Our dedicated support team ensuring smooth department operations.",
    labFacilities: "Lab Facilities",
    labMsg: "Modern practical equipment for hands-on learning.",
    syllabusMsg: "Download the updated curriculum for all semesters.",
    gallery: "Photo Gallery",
    footerMsg: "Empowering the next generation of Food Technology engineers with world-class education and practical skills.",
    quickLinks: "Quick Links",
    contact: "Contact Us",
    bteb: "BTEB Official",
    dte: "DTE Official",
    tmed: "TMED Official",
    cnpi: "CNPI Main Website",
    officeStaff: "Office Staff",
    location: "Barghoria, Chapainawabganj",
    rights: "All Rights Reserved.",
    searchPlaceholder: "Search students by name or roll...",
    exportCSV: "Export CSV",
    printList: "Print List",
    totalStudents: "Total Students",
    session: "Session",
    filterSemester: "All Semesters",
    filterShift: "All Shifts",
    filterGroup: "All Groups",
    developer: "Nasim Al Masud",
    studentPortalTitle: "Student Info Management",
    profile: "Profile",
    login: "Login",
    register: "Register",
    logout: "Logout",
    generateCV: "Generate CV",
    printCV: "Print CV",
    basicInfo: "Basic Info",
    eduInfo: "Education History",
    parentInfo: "Parental Info",
    addressInfo: "Addresses",
    proInfo: "Professional Info",
    finInfo: "Financial Info",
    verifyEmailMsg: "Please check your email for verification link.",
    updateProfileMsg: "Profile updated successfully.",
    adminPanel: "Admin Panel",
    bulkImport: "Bulk Import (Excel)",
    importSuccess: "Students imported successfully!",
    importError: "Error importing students.",
    downloadSample: "Download Sample Excel",
    semesterDistribution: "Student Distribution by Semester"
  },
  bn: {
    dept: "ফুড টেকনোলজি",
    inst: "চাঁপাইনবাবগঞ্জ পলিটেকনিক",
    home: "হোম",
    about: "আমাদের সম্পর্কে",
    teachers: "শিক্ষকবৃন্দ",
    staff: "অফিস ও ল্যাব স্টাফ",
    syllabus: "সিলেবাস",
    labs: "ল্যাব সমূহ",
    students: "শিক্ষার্থীবৃন্দ",
    studentPortal: "শিক্ষার্থী পোর্টাল",
    notices: "নোটিশ",
    heroTitle: "ফুড টেকনোলজি বিভাগ",
    heroSub: "শিখুন • উদ্ভাবন করুন • নেতৃত্ব দিন",
    meetTeachers: "শিক্ষকদের দেখুন",
    explore: "আরও জানুন",
    latestNews: "সর্বশেষ খবর",
    tickerMsg: "২০২৬ শিক্ষাবর্ষের ভর্তি চলছে। • আগামী সপ্তাহে ফুড ফেস্টিভ্যাল! • সেমিস্টার ফাইনালের রুটিন প্রকাশিত হয়েছে।",
    noticeBoard: "নোটিশ বোর্ড",
    allNotices: "সকল নোটিশ",
    noticeSearch: "নোটিশ খুঁজুন...",
    whyChoose: "কেন আমাদের নির্বাচন করবেন",
    modernLabs: "আধুনিক ল্যাব",
    expertTeachers: "অভিজ্ঞ শিক্ষক",
    careerSupport: "ক্যারিয়ার সহায়তা",
    aboutTitle: "ডিপার্টমেন্ট সম্পর্কে",
    aboutDesc: "চাঁপাইনবাবগঞ্জ পলিটেকনিক ইন্সটিটিউটের ফুড টেকনোলজি ডিপার্টমেন্ট বিশ্বমানের কারিগরি শিক্ষা প্রদানে নিবেদিত। বৈশ্বিক শিল্পের জন্য শিক্ষার্থীদের প্রস্তুত করতে আমরা খাদ্য প্রক্রিয়াজাতকরণ, সংরক্ষণ, মান নিয়ন্ত্রণ এবং আধুনিক ফুড ইঞ্জিনিয়ারিংয়ের উপর জোর দিয়ে থাকি।",
    historyTitle: "🏫 আমাদের ইতিহাস",
    historyDesc: "চাঁপাইনবাবগঞ্জ পলিটেকনিক ইনস্টিটিউটের ফুড টেকনোলজি বিভাগটি প্রতিষ্ঠিত হয় দেশের দ্রুত বিকাশমান খাদ্য শিল্পের জন্য উচ্চ দক্ষতা সম্পন্ন ও যোগ্য ডিপ্লোমা প্রকৌশলী তৈরি করার লক্ষ্যে। বর্তমান সময়ে দেশীয় ও বৈশ্বিক পর্যায়ে নিরাপদ, প্রক্রিয়াজাত ও মানসম্মত খাদ্যের চাহিদা ক্রমাগত বৃদ্ধি পাচ্ছে। এই চাহিদা পূরণের জন্য প্রযুক্তিগতভাবে দক্ষ জনবলের প্রয়োজনীয়তা বিবেচনা করেই এই বিভাগের সূচনা করা হয়।\n\nচাঁপাইনবাবগঞ্জ জেলা তার সমৃদ্ধ কৃষি সম্পদের জন্য বিশেষভাবে পরিচিত, বিশেষ করে আমসহ বিভিন্ন ফল ও ফসল উৎপাদনের জন্য। এই ভৌগোলিক সুবিধা খাদ্য প্রক্রিয়াজাতকরণ শিল্প গড়ে তোলার জন্য একটি বড় সম্ভাবনা তৈরি করে। এই স্থানীয় সম্পদকে সঠিকভাবে ব্যবহার করা এবং আধুনিক প্রযুক্তির মাধ্যমে মূল্য সংযোজন করার লক্ষ্যেই ফুড টেকনোলজি বিভাগটি প্রতিষ্ঠিত হয়, যা আঞ্চলিক ও জাতীয় উন্নয়নে গুরুত্বপূর্ণ ভূমিকা রাখে।\n\nপ্রতিষ্ঠার শুরুতে বিভাগটি সীমিত অবকাঠামো ও সুযোগ-সুবিধা নিয়ে যাত্রা শুরু করেছিল। তবে সময়ের সাথে সাথে ধারাবাহিক উন্নয়ন, শিক্ষা কর্তৃপক্ষের সহযোগিতা এবং অভিজ্ঞ শিক্ষকদের আন্তরিক প্রচেষ্টায় বিভাগটি একটি আধুনিক ও সুসংগঠিত বিভাগে পরিণত হয়েছে। বর্তমানে এখানে উন্নত ল্যাবরেটরি, আধুনিক যন্ত্রপাতি এবং ব্যবহারিকভিত্তিক শিক্ষার মাধ্যমে শিক্ষার্থীদের বাস্তবমুখী জ্ঞান প্রদান করা হয়, যা তাদের শিল্পক্ষেত্রে কাজ করার জন্য প্রস্তুত করে।\n\nবছরের পর বছর ধরে এই বিভাগ থেকে অনেক দক্ষ গ্র্যাজুয়েট বের হয়েছে, যারা দেশের বিভিন্ন নামকরা খাদ্য শিল্প প্রতিষ্ঠান এবং বিদেশেও সফলভাবে কাজ করছে। তাদের সাফল্য এই বিভাগের শিক্ষা ও প্রশিক্ষণের মানকে প্রতিফলিত করে। ফলে ফুড টেকনোলজি বিভাগটি এখন এই অঞ্চলের অন্যতম মর্যাদাপূর্ণ ও সম্ভাবনাময় বিভাগ হিসেবে পরিচিতি লাভ করেছে।\n\nএই বিভাগের ধারাবাহিক উন্নয়ন, উৎকর্ষের প্রতি অঙ্গীকার এবং ব্যবহারিক শিক্ষার উপর গুরুত্ব দেওয়ার কারণে এটি দেশের খাদ্য খাতে দক্ষ জনবল তৈরিতে গুরুত্বপূর্ণ ভূমিকা পালন করছে। এটি শুধু একটি শিক্ষা বিভাগ নয়, বরং ভবিষ্যতের জন্য দক্ষ কর্মশক্তি গড়ে তোলার একটি শক্তিশালী মাধ্যম।",
    missionTitle: "🎯 আমাদের লক্ষ্য",
    missionDesc: "আমাদের লক্ষ্য হলো এমন একটি সমন্বিত ও আধুনিক প্রযুক্তিগত শিক্ষা প্রদান করা, যার মাধ্যমে শিক্ষার্থীরা খাদ্য প্রক্রিয়াজাতকরণ (food processing), সংরক্ষণ (preservation) এবং মান নিয়ন্ত্রণ (quality control) বিষয়ে প্রয়োজনীয় দক্ষতা অর্জন করতে পারে এবং শিল্পক্ষেত্রের মানদণ্ড পূরণ করতে সক্ষম হয়।\n\nএই লক্ষ্য অর্জনের জন্য বিভাগটি শিক্ষার্থীদের তাত্ত্বিক জ্ঞানের পাশাপাশি ব্যবহারিক ও বাস্তবমুখী প্রশিক্ষণের উপর বিশেষ গুরুত্ব দেয়। শিক্ষার্থীরা আধুনিক ল্যাবরেটরি, উন্নত যন্ত্রপাতি এবং ইন্ডাস্ট্রি-ভিত্তিক শিক্ষার মাধ্যমে খাদ্য উৎপাদন ও প্রক্রিয়াজাতকরণের বিভিন্ন ধাপ সম্পর্কে গভীর ধারণা লাভ করে।\n\nএছাড়াও, তাদেরকে খাদ্য নিরাপত্তা, স্বাস্থ্যবিধি, আন্তর্জাতিক মান (যেমন ISO ও HACCP) এবং মান নিয়ন্ত্রণ পদ্ধতি সম্পর্কে প্রশিক্ষণ দেওয়া হয়, যাতে তারা বাস্তব কর্মক্ষেত্রে দক্ষতার সাথে কাজ করতে পারে।\n\nবিভাগটির আরেকটি গুরুত্বপূর্ণ উদ্দেশ্য হলো শিক্ষার্থীদের পেশাগত দক্ষতা, দায়িত্ববোধ এবং নৈতিক মূল্যবোধ গড়ে তোলা, যাতে তারা ভবিষ্যতে একজন সফল ফুড টেকনোলজিস্ট হিসেবে শিল্পক্ষেত্রে অবদান রাখতে পারে।\n\nসার্বিকভাবে, এই বিভাগের লক্ষ্য শুধু শিক্ষার্থীদের জ্ঞান প্রদান নয়, বরং তাদেরকে এমনভাবে গড়ে তোলা যাতে তারা আধুনিক খাদ্য শিল্পের চাহিদা পূরণে সক্ষম দক্ষ জনবল হিসেবে নিজেদের প্রতিষ্ঠিত করতে পারে।",
    visionTitle: "🌍 আমাদের রূপকল্প",
    visionDesc: "আমাদের রূপকল্প হলো ফুড টেকনোলজি শিক্ষায় একটি উৎকর্ষতার কেন্দ্র (Center of Excellence) হিসেবে প্রতিষ্ঠিত হওয়া, যা বৈশ্বিক খাদ্য খাতে উদ্ভাবন (innovation), মান উন্নয়ন এবং স্থায়িত্ব (sustainability) নিশ্চিত করতে গুরুত্বপূর্ণ ভূমিকা রাখবে।\n\nআমরা এমন একটি স্বীকৃত ও অগ্রণী শিক্ষা প্রতিষ্ঠান হিসেবে নিজেদের গড়ে তুলতে চাই, যেখানে আন্তর্জাতিক মানসম্পন্ন শিক্ষা ও প্রশিক্ষণ প্রদান করা হয়। আধুনিক প্রযুক্তি, গবেষণা এবং ব্যবহারিক শিক্ষার সমন্বয়ের মাধ্যমে শিক্ষার্থীদের এমনভাবে প্রস্তুত করা আমাদের লক্ষ্য, যাতে তারা দেশীয় ও আন্তর্জাতিক খাদ্য শিল্পে প্রতিযোগিতামূলকভাবে সফল হতে পারে।\n\nএই বিভাগ খাদ্য প্রক্রিয়াজাতকরণ, সংরক্ষণ এবং নতুন পণ্য উদ্ভাবনের ক্ষেত্রে সৃজনশীল চিন্তা ও গবেষণাকে উৎসাহিত করে। শিক্ষার্থীদের নতুন ধারণা তৈরি করতে এবং খাদ্যের মান, নিরাপত্তা ও উৎপাদন দক্ষতা উন্নত করতে উদ্বুদ্ধ করা হয়।\n\nআমাদের রূপকল্পের একটি গুরুত্বপূর্ণ দিক হলো স্থায়িত্ব নিশ্চিত করা। আমরা পরিবেশবান্ধব উৎপাদন পদ্ধতি, সম্পদের সঠিক ব্যবহার এবং খাদ্য অপচয় কমানোর উপর গুরুত্ব দিই, যাতে ভবিষ্যৎ প্রজন্মের জন্য নিরাপদ ও টেকসই খাদ্য ব্যবস্থা গড়ে তোলা যায়।\n\nনিরবচ্ছিন্ন উন্নয়ন, গবেষণা এবং শিল্প প্রতিষ্ঠানের সাথে সমন্বয়ের মাধ্যমে আমাদের লক্ষ্য দক্ষ ফুড টেকনোলজিস্ট তৈরি করা, যারা বৈশ্বিক খাদ্য খাতের উন্নয়নে গুরুত্বপূর্ণ ভূমিকা পালন করবে এবং ভবিষ্যতে খাদ্য নিরাপত্তা নিশ্চিত করতে সক্ষম হবে।",
    syllabusDesc: "ফুড টেকনোলজির জন্য বিটিইবি প্রবিধান ২০২২ সিলেবাস ডাউনলোড করুন।",
    academicInfo: "একাডেমিক তথ্য",
    facultyMsg: "খাদ্য উদ্ভাবনে অগ্রগামী আমাদের অসাধারণ শিক্ষক মণ্ডলীর সাথে পরিচিত হোন।",
    staffMsg: "বিভাগের কার্যক্রম সচল রাখতে আমাদের নিবেদিত অফিস ও ল্যাব কর্মী বৃন্দ।",
    labFacilities: "ল্যাব সুবিধা",
    labMsg: "হাতে-কলমে শেখার জন্য আধুনিক ব্যবহারিক যন্ত্রাংশ।",
    syllabusMsg: "সকল সেমিস্টারের আপডেটেড সিলেবাস ডাউনলোড করুন।",
    gallery: "ফটো গ্যালারি",
    footerMsg: "বিশ্বমানের শিক্ষা ও দক্ষ ফুড টেকনোলজি ইঞ্জিনিয়ার গড়ে তোলাই আমাদের লক্ষ্য।",
    quickLinks: "প্রয়োজনীয় লিংক",
    contact: "যোগাযোগ",
    bteb: "কারিগরি শিক্ষা বোর্ড",
    dte: "কারিগরি শিক্ষা অধিদপ্তর",
    tmed: "কারিগরি শিক্ষা ও মাদ্রাসা বিভাগ",
    cnpi: "সিএনপিআই মূল ওয়েবসাইট",
    officeStaff: "অফিস স্টাফ",
    location: "বারঘরিয়া, চাঁপাইনবাবগঞ্জ",
    rights: "সর্বস্বত্ব সংরক্ষিত।",
    searchPlaceholder: "নাম বা রোলে শিক্ষার্থী খুঁজুন...",
    exportCSV: "সিএসভি এক্সপোর্ট",
    printList: "তালিকা প্রিন্ট",
    totalStudents: "মোট শিক্ষার্থী",
    session: "সেশন",
    filterSemester: "সকল সেমিস্টার",
    filterShift: "সকল শিফট",
    filterGroup: "সকল গ্রুপ",
    developer: "নাসিম আল মাসুদ",
    studentPortalTitle: "শিক্ষার্থী তথ্য ব্যবস্থাপনা",
    profile: "প্রোফাইল",
    login: "লগইন",
    register: "রেজিস্ট্রেশন",
    logout: "লগআউট",
    generateCV: "সিভি জেনারেট",
    printCV: "সিভি প্রিন্ট",
    basicInfo: "বেসিক তথ্য",
    eduInfo: "শিক্ষাগত ইতিহাস",
    parentInfo: "পিতামাতার তথ্য",
    addressInfo: "ঠিকানা",
    proInfo: "পেশাগত তথ্য",
    finInfo: "আর্থিক তথ্য",
    verifyEmailMsg: "ভেরিফিকেশন লিঙ্কের জন্য আপনার ইমেল চেক করুন।",
    updateProfileMsg: "প্রোফাইল সফলভাবে আপডেট করা হয়েছে।",
    adminPanel: "এডমিন প্যানেল",
    bulkImport: "বাল্ক ইমপোর্ট (এক্সেল)",
    importSuccess: "শিক্ষার্থীরা সফলভাবে ইমপোর্ট করা হয়েছে!",
    importError: "ইমপোর্ট করতে সমস্যা হয়েছে।",
    downloadSample: "স্যাম্পল এক্সেল ডাউনলোড করুন",
    semesterDistribution: "সেমিস্টার অনুযায়ী শিক্ষার্থী বিভাজন"
  }
};

export default function App() {
  const [lang, setLang] = useState<Language>(() => (localStorage.getItem('lang') as Language) || 'bn');
  const [activeTab, setActiveTab] = useState('home');
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const [loading, setLoading] = useState(true);
  const [user, setUser] = useState<User | null>(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [studentProfile, setStudentProfile] = useState<StudentProfile | null>(null);
  const [authMode, setAuthMode] = useState<'login' | 'register'>('login');
  const [authError, setAuthError] = useState('');
  const [authLoading, setAuthLoading] = useState(false);
  
  const [teachers, setTeachers] = useState<Teacher[]>([]);
  const [staff, setStaff] = useState<Teacher[]>([]);
  const [notices, setNotices] = useState<Notice[]>([]);
  const [labs, setLabs] = useState<Lab[]>([]);
  const [students, setStudents] = useState<Student[]>([]);
  const [registeredStudents, setRegisteredStudents] = useState<Student[]>([]);
  const [semesters, setSemesters] = useState<Record<string, any[]>>({});

  const [isAddingNotice, setIsAddingNotice] = useState(false);
  const [newNotice, setNewNotice] = useState({ text_en: '', text_bn: '', file: '' });

  const [searchQuery, setSearchQuery] = useState('');
  const [semesterFilter, setSemesterFilter] = useState('');
  const [shiftFilter, setShiftFilter] = useState('');
  const [groupFilter, setGroupFilter] = useState('');
  const [teacherSearch, setTeacherSearch] = useState('');
  const [noticeSearch, setNoticeSearch] = useState('');
  const [teacherFilter, setTeacherFilter] = useState('All');
  const [selectedCourse, setSelectedCourse] = useState<any>(null);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 8;

  const [showBackToTop, setShowBackToTop] = useState(false);

  const t = translations[lang];

  const dynamicTicker = useMemo(() => {
    if (notices.length === 0) return t.tickerMsg;
    return notices.slice(0, 10).map(n => lang === 'en' ? n.text_en : n.text_bn).join(' • ');
  }, [notices, lang, t.tickerMsg]);

  useEffect(() => {
    const handleScroll = () => {
      setShowBackToTop(window.scrollY > 400);
    };
    window.addEventListener('scroll', handleScroll);
    return () => window.removeEventListener('scroll', handleScroll);
  }, []);

  useEffect(() => {
    localStorage.setItem('lang', lang);
  }, [lang]);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setIsAdmin(currentUser?.email === 'mimmaevent2003@gmail.com');
      if (currentUser) {
        fetchProfile(currentUser.uid);
      } else {
        setStudentProfile(null);
      }
    });
    
    // Test connection
    import('./lib/firebase').then(m => m.testConnection());

    return () => unsubscribe();
  }, []);

  const fetchProfile = async (uid: string) => {
    try {
      const { getDoc, doc } = await import('firebase/firestore');
      const docRef = doc(db, 'student_profiles', uid);
      const docSnap = await getDoc(docRef);
      if (docSnap.exists()) {
        setStudentProfile(docSnap.data() as StudentProfile);
      }
    } catch (error) {
      console.error("Fetch profile error:", error);
    }
  };

  const handleAuth = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setAuthError('');
    setAuthLoading(true);
    const formData = new FormData(e.currentTarget);
    const email = formData.get('email') as string;
    const password = formData.get('password') as string;
    const name = formData.get('name') as string;
    const roll = formData.get('roll') as string;
    const reg = formData.get('registration') as string;
    const mobile = formData.get('mobile') as string;
    const semester = formData.get('semester') as string;
    const shift = formData.get('shift') as string;
    const group = formData.get('group') as string;
    const session = formData.get('session') as string;
    const image = formData.get('image') as string;

    try {
      if (authMode === 'register') {
        const userCredential = await createUserWithEmailAndPassword(auth, email, password);
        await sendEmailVerification(userCredential.user);
        await updateProfile(userCredential.user, { displayName: name, photoURL: image });
        
        // Create initial profile
        const initialProfile: StudentProfile = {
          userId: userCredential.user.uid,
          basic: { name, roll, registration: reg, email, mobile, semester, shift, group, session, image },
          education: [],
          parents: {
            father: { name: '', occupation: '', mobile: '', nid: '', dob: '' },
            mother: { name: '', occupation: '', mobile: '', nid: '', dob: '' },
            guardian: { name: '', relation: '', mobile: '', occupation: '' }
          },
          address: { current: '', permanent: '' },
          professional: { objective: '', experience: '', certificates: '', languages: '' },
          financial: { bankName: '', accName: '', accNo: '', accNid: '', bkash: '', rocket: '' }
        };
        
        const { setDoc, doc } = await import('firebase/firestore');
        await setDoc(doc(db, 'student_profiles', userCredential.user.uid), initialProfile);
        setStudentProfile(initialProfile);
        alert(t.verifyEmailMsg);
      } else {
        await signInWithEmailAndPassword(auth, email, password);
      }
    } catch (error: any) {
      if (error.code === 'auth/operation-not-allowed') {
        setAuthError(lang === 'en' 
          ? "Registration is currently disabled. Please enable Email/Password at Firebase Console." 
          : "রেজিষ্ট্রেশন বর্তমানে বন্ধ রয়েছে। অনুগ্রহ করে ফায়ারবেস কনসোল থেকে Email/Password অপশনটি ইনাবল করুন।");
      } else {
        setAuthError(error.message);
      }
    } finally {
      setAuthLoading(false);
    }
  };

  const saveProfile = async (updatedProfile: StudentProfile) => {
    if (!user) return;
    try {
      const { setDoc, doc } = await import('firebase/firestore');
      await setDoc(doc(db, 'student_profiles', user.uid), updatedProfile);
      setStudentProfile(updatedProfile);
      alert(t.updateProfileMsg);
    } catch (error) {
      console.error("Save profile error:", error);
      alert("Error saving profile.");
    }
  };

  const handleBulkImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !isAdmin) return;

    try {
      const reader = new FileReader();
      reader.onload = async (evt) => {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];

        const { setDoc, doc, getDocs, query, collection, where } = await import('firebase/firestore');
        const importPromises = data.map(async (item) => {
          const studentData: Student = {
            name: String(item.name || ''),
            roll: String(item.roll) || '',
            registration: String(item.registration) || '',
            semester: String(item.semester || ''),
            shift: String(item.shift || ''),
            group: String(item.group || ''),
            session: String(item.session || ''),
            mobile: String(item.mobile) || '',
            image: String(item.image || '')
          };
          
          if (!studentData.roll) return;

          // 1. Update master directory
          await setDoc(doc(db, 'students', studentData.roll), studentData);

          // 2. Sync with portal profile if exists
          const profileQuery = query(collection(db, 'student_profiles'), where('basic.roll', '==', studentData.roll));
          const profileSnap = await getDocs(profileQuery);
          
          if (!profileSnap.empty) {
            const profileDoc = profileSnap.docs[0];
            const currentProfile = profileDoc.data() as StudentProfile;
            
            // Sync basic info but preserve portal-specific fields if they have data
            await updateDoc(doc(db, 'student_profiles', profileDoc.id), {
              'basic.name': studentData.name,
              'basic.registration': studentData.registration,
              'basic.semester': studentData.semester,
              'basic.shift': studentData.shift,
              'basic.group': studentData.group,
              'basic.session': studentData.session,
              'basic.mobile': studentData.mobile,
              // Only update image if item.image is provided
              ...(item.image ? { 'basic.image': item.image } : {})
            });
          }
        });

        await Promise.all(importPromises);
        alert(t.importSuccess);
      };
      reader.readAsBinaryString(file);
    } catch (error) {
      console.error("Bulk import error:", error);
      alert(t.importError);
    }
  };

  const downloadSampleExcel = () => {
    const sampleData = [
      {
        name: "John Doe",
        roll: "123456",
        registration: "987654321",
        semester: "5th",
        shift: "1st",
        group: "A",
        session: "2023-24",
        mobile: "01700000000",
        image: "https://example.com/photo.jpg"
      },
      {
        name: "Jane Smith",
        roll: "112233",
        registration: "556677889",
        semester: "3rd",
        shift: "2nd",
        group: "B",
        session: "2022-23",
        mobile: "01800000000",
        image: ""
      }
    ];

    const ws = XLSX.utils.json_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Students");
    XLSX.writeFile(wb, "Student_Import_Sample.xlsx");
  };

  const login = async () => {
    try {
      await signInWithPopup(auth, googleProvider);
    } catch (error) {
      console.error("Login error:", error);
    }
  };

  const logout = () => signOut(auth);

  const addNotice = async () => {
    if (!isAdmin) return;
    try {
      await addDoc(collection(db, 'notices'), {
        ...newNotice,
        date: new Date().toISOString().split('T')[0],
        createdAt: serverTimestamp()
      });
      setIsAddingNotice(false);
      setNewNotice({ text_en: '', text_bn: '', file: '' });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'notices');
    }
  };

  useEffect(() => {
    // Basic static data
    fetch('/semester.json').then(res => res.ok && res.json()).then(data => data && setSemesters(data));

    // Public Firebase listeners
    const unsubscribeNotices = onSnapshot(
      query(collection(db, 'notices'), orderBy('date', 'desc')), 
      (snapshot) => {
        const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as any as Notice));
        if (data.length > 0) setNotices(data);
        else fetch('/notices.json').then(res => res.json()).then(setNotices);
      },
      (error) => console.warn("Notices listener info:", error.message)
    );

    const unsubscribePeople = onSnapshot(
      collection(db, 'people'),
      (snapshot) => {
        const allPeople = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as any as Teacher & { type: string }));
        if (allPeople.length > 0) {
          setTeachers(allPeople.filter(p => p.type === 'teacher'));
          setStaff(allPeople.filter(p => p.type === 'staff'));
        } else {
          fetch('/teachers.json').then(res => res.json()).then(setTeachers);
          fetch('/staff.json').then(res => res.json()).then(setStaff);
        }
      },
      (error) => console.warn("People listener info:", error.message)
    );

    const unsubscribeLabs = onSnapshot(
      collection(db, 'labs'),
      (snapshot) => {
        const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as any as Lab));
        if (data.length > 0) setLabs(data);
        else fetch('/labs.json').then(res => res.json()).then(setLabs);
      },
      (error) => console.warn("Labs listener info:", error.message)
    );

    setLoading(false);

    return () => {
      unsubscribeNotices();
      unsubscribePeople();
      unsubscribeLabs();
    };
  }, []);

  // Student listeners
  useEffect(() => {
    if (!user) {
      // Fallback to static if not admin or not logged in
      if (students.length === 0) {
        fetch('/students.json').then(res => res.json()).then(setStudents);
      }
      setRegisteredStudents([]);
      return;
    }

    const unsubscribeStudents = onSnapshot(
      collection(db, 'students'),
      (snapshot) => {
        const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as any as Student));
        if (data.length > 0) setStudents(data);
      },
      (error) => console.warn("Students listener info:", error.message)
    );

    const unsubscribeRegistered = onSnapshot(
      collection(db, 'student_profiles'),
      (snapshot) => {
        const data = snapshot.docs.map(doc => {
          const profile = doc.data() as StudentProfile;
          return {
            name: profile.basic.name,
            roll: profile.basic.roll,
            registration: profile.basic.registration,
            semester: profile.basic.semester || 'N/A',
            shift: profile.basic.shift || 'N/A',
            group: profile.basic.group || 'N/A',
            mobile: profile.basic.mobile,
            image: profile.basic.image || ''
          } as Student;
        });
        setRegisteredStudents(data);
      },
      (error) => console.warn("Registered students listener info:", error.message)
    );

    return () => {
      unsubscribeStudents();
      unsubscribeRegistered();
    };
  }, [user]);

  const allAvailableStudents = useMemo(() => {
    // Precedence: Portal Profiles (registeredStudents) > Bulk Data (students)
    const combined = [...registeredStudents];
    
    // Add students from bulk list if NOT already in portal list
    students.forEach(s => {
      if (!combined.find(rs => rs.roll === s.roll)) {
        combined.push(s);
      }
    });

    return combined;
  }, [students, registeredStudents]);

  const semesterStats = useMemo(() => {
    const semList = ['1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th'];
    return semList.map(sem => ({
      name: sem,
      count: registeredStudents.filter(s => s.semester === sem).length
    }));
  }, [registeredStudents]);

  const filteredTeachers = useMemo(() => {
    const data = activeTab === 'teachers' ? teachers : staff;
    const search = teacherSearch.toLowerCase().trim();
    
    return data.filter(person => {
      const matchesSearch = !search || 
        person.name_en.toLowerCase().includes(search) || 
        person.name_bn.includes(search) ||
        person.position_en.toLowerCase().includes(search) ||
        person.position_bn.includes(search);
      
      let matchesFilter = teacherFilter === 'All';
      if (!matchesFilter) {
        if (teacherFilter === 'Instructor') {
          // Strictly match Instructor but exclude Junior and Chief Instructor
          matchesFilter = person.position_en.includes('Instructor') && 
                         !person.position_en.includes('Junior') && 
                         !person.position_en.includes('Chief');
        } else {
          matchesFilter = person.position_en.includes(teacherFilter);
        }
      }
      
      return matchesSearch && matchesFilter;
    });
  }, [activeTab, teachers, staff, teacherSearch, teacherFilter]);

  const filteredStudents = useMemo(() => {
    return allAvailableStudents.filter(s => {
      const matchesSearch = s.name.toLowerCase().includes(searchQuery.toLowerCase()) || 
                          s.roll.includes(searchQuery);
      const matchesSemester = !semesterFilter || s.semester === semesterFilter;
      const matchesShift = !shiftFilter || s.shift === shiftFilter;
      const matchesGroup = !groupFilter || s.group === groupFilter;
      return matchesSearch && matchesSemester && matchesShift && matchesGroup;
    });
  }, [allAvailableStudents, searchQuery, semesterFilter, shiftFilter, groupFilter]);

  const paginatedStudents = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return filteredStudents.slice(start, start + itemsPerPage);
  }, [filteredStudents, currentPage]);

  const totalPages = Math.ceil(filteredStudents.length / itemsPerPage);

  const filteredAllNotices = useMemo(() => {
    return notices.filter(n => 
      n.text_en.toLowerCase().includes(noticeSearch.toLowerCase()) || 
      n.text_bn.includes(noticeSearch) ||
      n.date.includes(noticeSearch)
    );
  }, [notices, noticeSearch]);

  const exportCSV = () => {
    let csv = "Name,Roll,Registration,Semester,Shift,Group,Session,Mobile,Guardian\n";
    filteredStudents.forEach(s => {
      csv += `"${s.name}","${s.roll}","${s.registration}","${s.semester}","${s.shift}","${s.group}","${s.session || ''}","${s.mobile}","${s.guardian || ''}"\n`;
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.setAttribute("download", "student_list.csv");
    link.click();
  };

  const printList = () => {
    const printWindow = window.open('', '', 'width=1000,height=700');
    if (!printWindow) return;
    const printTime = new Date().toLocaleString();
    let tableRows = filteredStudents.map((s, i) => `
        <tr>
            <td>${i + 1}</td>
            <td>${s.name}</td>
            <td>${s.roll}</td>
            <td>${s.registration}</td>
            <td>${s.semester}</td>
            <td>${s.group}</td>
            <td>${s.shift}</td>
            <td>${s.session || '-'}</td>
            <td>${s.mobile}</td>
            <td>${s.guardian || '-'}</td>
        </tr>
    `).join('');

    printWindow.document.write(`
    <html>
    <head>
        <title>Student List - CNPI</title>
        <style>
            @page { size: A4 portrait; margin: 20mm 15mm; }
            body { font-family: 'Segoe UI', Arial, sans-serif; color: #333; margin: 0; padding: 0; }
            .header-table { width: 100%; border: none; margin-bottom: 20px; border-bottom: 2px solid #116530; padding-bottom: 10px; }
            .header-table td { border: none; text-align: center; vertical-align: middle; }
            .header-table img { width: 65px; height: auto; }
            .header-text h2 { margin: 0; color: #116530; font-size: 20px; text-transform: uppercase; }
            .header-text p { margin: 2px 0; font-size: 13px; font-weight: bold; }
            table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 11px; }
            th, td { border: 1px solid #000; padding: 6px 4px; text-align: center; }
            th { background-color: #f2f2f2 !important; -webkit-print-color-adjust: exact; }
            thead { display: table-header-group; } 
            tr { page-break-inside: avoid; }
            .signature-section { margin-top: 60px; display: flex; justify-content: space-between; }
            .sign-box { width: 22%; text-align: center; font-size: 11px; font-weight: bold; }
            .line { border-top: 1px solid #000; margin-bottom: 5px; }
            .print-footer { position: fixed; bottom: 0; width: 100%; font-size: 9px; display: flex; justify-content: space-between; border-top: 1px solid #ddd; padding-top: 5px; }
        </style>
    </head>
    <body onload="window.print()">
        <table class="header-table">
            <tr>
                <td style="width: 15%;"><img src="https://upload.wikimedia.org/wikipedia/commons/thumb/8/84/Government_Seal_of_Bangladesh.svg/500px-Government_Seal_of_Bangladesh.svg.png"></td>
                <td style="width: 70%;" class="header-text">
                    <h2>Chapainawabganj Polytechnic Institute</h2>
                    <p>Department of Food Technology</p>
                    <p style="background: #eee; display: inline-block; padding: 2px 10px; border-radius: 4px;">Student Information Record</p>
                </td>
                <td style="width: 15%;"><img src="https://objectstorage.ap-dcc-gazipur-1.oraclecloud15.com/n/axvjbnqprylg/b/V2Ministry/o/office-cnpi-chapainawabganj/2024/12/d24884ba85bd4a638c87214b8f286ea1.png"></td>
            </tr>
        </table>
        <table>
            <thead>
                <tr>
                    <th>SL</th><th>Name</th><th>Roll</th><th>Registration</th><th>Sem</th><th>Group</th><th>Shift</th><th>Session</th><th>Mobile</th><th>Guardian</th>
                </tr>
            </thead>
            <tbody>${tableRows}</tbody>
        </table>
        <div class="signature-section">
            <div class="sign-box"><div class="line"></div>Class Teacher</div>
            <div class="sign-box"><div class="line"></div>Head of Dept.</div>
            <div class="sign-box"><div class="line"></div>Vice Principal</div>
            <div class="sign-box"><div class="line"></div>Principal</div>
        </div>
        <div class="print-footer">
            <div>Printed on: ${printTime}</div>
            <div>Generated by: FoodTech CNPI Portal</div>
        </div>
    </body>
    </html>
    `);
    printWindow.document.close();
  };

  if (loading) {
    return (
      <div className="fixed inset-0 bg-brand-green flex items-center justify-center z-[100]">
        <motion.div
           animate={{ rotate: 360, scale: [1, 1.1, 1] }}
           transition={{ duration: 0.8, repeat: Infinity }}
        >
          <Leaf className="w-16 h-16 text-white" />
        </motion.div>
      </div>
    );
  }

  return (
    <div className={`min-h-screen font-sans bg-[#fbfcfd] flex flex-col ${lang === 'bn' ? 'bn' : ''}`}>
      <AnimatePresence>
        {isMenuOpen && (
          <>
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsMenuOpen(false)}
              className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] lg:hidden"
            />
            <motion.div 
              initial={{ x: '100%' }}
              animate={{ x: 0 }}
              exit={{ x: '100%' }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="fixed top-0 right-0 bottom-0 w-[85%] max-w-xs bg-white z-[101] lg:hidden shadow-2xl flex flex-col"
            >
              <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                <div className="flex items-center gap-3">
                  <div className="bg-brand-gradient p-2 rounded-xl text-white shadow-lg shadow-brand-green/20">
                    <Leaf size={18} />
                  </div>
                  <div className="flex flex-col">
                    <span className="font-black text-sm tracking-tight text-slate-900 uppercase">Department</span>
                    <span className="text-[9px] font-bold text-brand-green tracking-[0.2em] uppercase">Food Tech</span>
                  </div>
                </div>
                <button 
                  onClick={() => setIsMenuOpen(false)}
                  className="p-2.5 bg-white border border-slate-200 text-slate-400 hover:text-brand-green rounded-xl transition-all active:scale-95 shadow-sm"
                >
                  <X size={18} />
                </button>
              </div>
              
              <div className="flex-grow p-5 space-y-2 overflow-y-auto bg-white">
                {[
                  { id: 'home', label: t.home, icon: Home },
                  { id: 'about', label: t.about, icon: Info },
                  { id: 'teachers', label: t.teachers, icon: Users },
                  { id: 'syllabus', label: t.syllabus, icon: BookOpen },
                  { id: 'students', label: t.students, icon: UserIcon },
                  { id: 'notices', label: t.notices, icon: Bell },
                  { id: 'contact', label: t.contact, icon: Mail },
                  { id: 'student-portal', label: 'Access Portal', icon: Lock }
                ].map(item => (
                  <button
                    key={item.id}
                    onClick={() => { setActiveTab(item.id); setIsMenuOpen(false); }}
                    className={`mobile-menu-item w-full ${activeTab === item.id ? 'bg-brand-accent text-brand-green shadow-sm' : ''} group`}
                  >
                    <div className={`p-2 rounded-lg transition-colors ${activeTab === item.id ? 'bg-brand-green text-white' : 'bg-slate-50 text-slate-400 group-hover:bg-brand-green/10 group-hover:text-brand-green'}`}>
                      <item.icon size={16} />
                    </div>
                    {item.label}
                  </button>
                ))}
              </div>

              <div className="p-6 border-t border-slate-100 bg-slate-50/50">
                <button 
                  onClick={() => { setLang(l => l === 'en' ? 'bn' : 'en'); setIsMenuOpen(false); }}
                  className="w-full flex items-center justify-between p-4 rounded-2xl bg-white border border-slate-200 text-slate-600 font-bold text-[10px] uppercase tracking-widest shadow-sm transition-all active:scale-95"
                >
                  <div className="flex items-center gap-3">
                    <Globe size={14} className="text-brand-green" />
                    {lang === 'en' ? 'বাংলা সংস্করণ' : 'English View'}
                  </div>
                  <ChevronRight size={14} className="text-slate-300" />
                </button>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>
              
              <div className="p-6 border-t border-gray-100 italic text-center">
                <p className="text-[10px] text-gray-400 font-bold uppercase tracking-widest leading-relaxed">
                  Department of Food Technology<br/>CNPI Portal v2.0
                </p>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {selectedCourse && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setSelectedCourse(null)}
              className="absolute inset-0 bg-slate-900/80 backdrop-blur-md"
            />
            <motion.div 
              initial={{ scale: 0.9, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.9, opacity: 0, y: 20 }}
              className="relative bg-white w-full max-w-2xl rounded-[3rem] shadow-2xl overflow-hidden border border-brand-green/10"
            >
              <div className="bg-brand-green p-10 text-white relative">
                <Leaf className="absolute top-10 right-10 text-white/10 w-32 h-32 rotate-12" />
                <button 
                  onClick={() => setSelectedCourse(null)}
                  className="absolute top-8 right-8 p-3 hover:bg-white/10 rounded-2xl transition-colors"
                >
                  <X size={24} />
                </button>
                <div className="flex flex-col gap-4">
                   <div className="inline-block px-4 py-1.5 bg-brand-yellow text-brand-green text-[10px] font-black uppercase tracking-widest rounded-full w-fit">
                     Course Overview
                   </div>
                   <h2 className="text-3xl md:text-5xl font-black tracking-tighter leading-none">{selectedCourse.name}</h2>
                   <p className="text-sm font-black text-white/60 tracking-widest uppercase italic">
                     {selectedCourse.semester}{selectedCourse.semester === 1 ? 'st' : selectedCourse.semester === 2 ? 'nd' : selectedCourse.semester === 3 ? 'rd' : 'th'} Semester • Food Technology
                   </p>
                </div>
              </div>

              <div className="p-10 space-y-10">
                <div className="grid grid-cols-2 gap-6">
                   <div className="bg-gray-50 p-6 rounded-3xl border border-gray-100">
                      <p className="text-[10px] text-gray-400 font-black uppercase tracking-widest mb-2">Academic Code</p>
                      <p className="text-xl font-black text-slate-800">{selectedCourse.code || 'FT-2026'}</p>
                   </div>
                   <div className="bg-gray-50 p-6 rounded-3xl border border-gray-100">
                      <p className="text-[10px] text-gray-400 font-black uppercase tracking-widest mb-2">Subject Rating</p>
                      <div className="flex gap-1 text-brand-yellow">
                         {[1, 2, 3, 4, 5].map(i => <Globe key={i} size={14} fill="currentColor" />)}
                      </div>
                   </div>
                </div>

                <div className="space-y-4">
                   <h4 className="text-sm font-black text-slate-400 uppercase tracking-widest flex items-center gap-3">
                     <div className="h-0.5 w-8 bg-brand-green"></div> What You'll Learn
                   </h4>
                   <p className="text-lg text-slate-600 font-medium leading-relaxed italic border-l-4 border-brand-yellow pl-6">
                     This course provides comprehensive insights into {selectedCourse.name.toLowerCase()} for processing, quality assurance, and industrial application as per BTEB Regulation 2022.
                   </p>
                </div>

                <div className="flex flex-col sm:flex-row gap-6 pt-6">
                   <a 
                     href={selectedCourse.pdf} 
                     target="_blank" 
                     rel="noopener noreferrer"
                     className="flex-1 bg-brand-green text-white px-8 py-5 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-brand-green/90 transition-all shadow-xl shadow-brand-green/20 text-center flex items-center justify-center gap-3"
                   >
                     <Download size={18} /> Official Syllabus
                   </a>
                   <button 
                     onClick={() => setSelectedCourse(null)}
                     className="px-8 py-5 rounded-2xl font-black text-xs uppercase tracking-widest border-2 border-slate-100 text-slate-400 hover:border-brand-green hover:text-brand-green transition-all"
                   >
                     Close Portal
                   </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Top Banner */}
      <div className="bg-brand-green text-white text-[10px] md:text-xs py-2 px-4 shadow-inner relative z-50">
        <div className="max-w-7xl mx-auto flex justify-between items-center">
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-1.5 border-r border-brand-green/20 pr-4">
              <Clock size={12} className="shrink-0" />
              <span>{new Date().toLocaleTimeString(lang === 'bn' ? 'bn-BD' : 'en-US')}</span>
            </div>
            <span className="hidden sm:inline font-medium opacity-90">{new Date().toLocaleDateString(lang === 'bn' ? 'bn-BD' : 'en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</span>
          </div>
          <button 
            onClick={() => setLang(l => l === 'en' ? 'bn' : 'en')}
            className="flex items-center gap-2 hover:bg-brand-green/20 px-3 py-1 rounded-full transition-all font-black text-[10px] uppercase tracking-widest border border-white/20"
          >
            <Globe size={11} />
            {lang === 'en' ? 'বাংলা সংস্করণ' : 'English View'}
          </button>
        </div>
      </div>

      {/* Navigation */}
      <header className="sticky top-0 z-50 bg-white/90 backdrop-blur-xl border-b border-slate-100 shadow-sm">
        <nav className="max-w-7xl mx-auto px-4 h-16 md:h-20 flex items-center justify-between">
          <a href="#" onClick={() => setActiveTab('home')} className="flex items-center gap-2 md:gap-4 group">
            <div className="bg-brand-gradient p-1.5 md:p-2.5 rounded-lg md:rounded-2xl text-white shadow-lg shadow-brand-green/20 group-hover:scale-110 transition-all duration-500">
              <Leaf size={18} className="md:w-6 md:h-6" />
            </div>
            <div className="flex flex-col">
              <h1 className="text-xs md:text-xl font-bold md:font-black text-slate-900 leading-none tracking-tight font-display uppercase">
                {lang === 'en' ? 'Food Tech Dept' : 'ফুড টেকনোলজি বিভাগ'}
              </h1>
              <p className="text-[7px] md:text-[10px] text-brand-green font-bold tracking-widest uppercase mt-0.5 opacity-80">{t.inst}</p>
            </div>
          </a>

          <div className="hidden lg:flex items-center gap-6">
            {[
              { id: 'home', label: t.home },
              { id: 'about', label: t.about },
              { id: 'teachers', label: t.teachers },
              { id: 'syllabus', label: t.syllabus },
              { id: 'students', label: t.students },
              { id: 'notices', label: t.notices },
              { id: 'contact', label: t.contact }
            ].map(tab => (
              <button
                key={tab.id}
                onClick={() => setActiveTab(tab.id)}
                className={`nav-link ${activeTab === tab.id ? 'nav-link-active' : ''}`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          <div className="flex items-center gap-2 md:gap-4">
            <button 
              onClick={() => setActiveTab('student-portal')}
              className="bg-brand-gradient text-white px-3 md:px-6 py-2 md:py-3 rounded-full text-[9px] md:text-[10px] font-black uppercase tracking-widest transition-all shadow-lg shadow-brand-green/20 active:scale-95 flex items-center gap-2"
            >
              <UserIcon size={14} className="md:w-4 md:h-4" /> portal
            </button>

            <button 
              className="lg:hidden p-2 text-slate-500 hover:text-brand-green transition-colors"
              onClick={() => setIsMenuOpen(true)}
            >
              <Menu size={20} />
            </button>
          </div>
        </nav>
      </header>

      {/* Main Content */}
      <main className="flex-grow">
        <AnimatePresence mode="wait">
          <motion.div
            key={activeTab}
            initial={{ opacity: 0, y: 10 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -10 }}
            transition={{ duration: 0.3 }}
          >
            {activeTab === 'home' && (
              <div className="pb-16 md:pb-24">
                {/* Hero */}
                <section className="relative h-[65vh] md:h-[75vh] flex items-center justify-start overflow-hidden bg-slate-900 px-6 md:px-12">
                  <div className="absolute inset-0 z-0">
                    <img 
                       src="https://images.unsplash.com/photo-1550505393-fa79c2336ec8?auto=format&fit=crop&q=80" 
                       className="w-full h-full object-cover opacity-50 animate-slow-zoom"
                       alt="Background"
                    />
                    <div className="absolute inset-0 bg-gradient-to-r from-brand-green/90 via-brand-green/30 to-transparent"></div>
                  </div>
                  <div className="relative z-10 max-w-4xl mobile-center">
                    <motion.div 
                       initial={{ x: -20, opacity: 0 }}
                       whileInView={{ x: 0, opacity: 1 }}
                       transition={{ duration: 0.6 }}
                       className="inline-flex items-center gap-2 bg-brand-yellow text-brand-green px-4 py-1.5 rounded-full text-[8px] md:text-[9px] font-black uppercase tracking-[0.25em] mb-4 shadow-xl"
                    >
                      <Pin size={10} className="rotate-45" /> Food Science Excellence
                    </motion.div>
                    <motion.h1 
                       initial={{ y: 30, opacity: 0 }}
                       whileInView={{ y: 0, opacity: 1 }}
                       transition={{ delay: 0.1, duration: 0.8 }}
                       className="text-4xl md:text-8xl font-black text-white mb-6 leading-[0.95] tracking-tight font-display"
                    >
                      {lang === 'en' ? 'Shape the Future of Food' : 'খাদ্য প্রযুক্তির ভবিষ্যৎ গড়ুন'}
                    </motion.h1>
                    <motion.p 
                       initial={{ y: 30, opacity: 0 }}
                       whileInView={{ y: 0, opacity: 1 }}
                       transition={{ delay: 0.2, duration: 0.8 }}
                       className="text-sm md:text-lg text-white/90 font-medium mb-8 md:mb-10 tracking-normal max-w-xl leading-relaxed"
                    >
                      {t.aboutDesc}
                    </motion.p>
                    <motion.div 
                       initial={{ y: 30, opacity: 0 }}
                       whileInView={{ y: 0, opacity: 1 }}
                       transition={{ delay: 0.3, duration: 0.8 }}
                       className="flex flex-wrap justify-center md:justify-start gap-4"
                    >
                      <button 
                         onClick={() => setActiveTab('notices')}
                         className="bg-brand-gradient hover:opacity-90 text-white px-6 md:px-8 py-3 md:py-4 rounded-full font-black text-[10px] md:text-[11px] uppercase tracking-widest transition-all shadow-xl shadow-brand-green/30 active:scale-95 flex items-center gap-2 md:gap-3"
                      >
                        {lang === 'en' ? 'Admission' : 'ভর্তি চলছে'} <ChevronRight size={16} />
                      </button>
                      <button 
                         onClick={() => setActiveTab('about')}
                         className="bg-white/10 hover:bg-white/20 backdrop-blur-md text-white border border-white/30 px-6 md:px-8 py-3 md:py-4 rounded-full font-black text-[10px] md:text-[11px] uppercase tracking-widest transition-all active:scale-95"
                      >
                        {t.explore}
                      </button>
                    </motion.div>
                  </div>
                </section>

                {/* News Bar */}
                <div className="bg-white border-y border-slate-100 h-14 md:h-16 relative overflow-hidden flex items-center shadow-md">
                  <div className="absolute left-0 top-0 bottom-0 bg-brand-gradient text-white px-6 md:px-10 font-bold text-[9px] md:text-[10px] uppercase tracking-widest z-20 flex items-center gap-3 shadow-xl">
                    <div className="w-2 h-2 bg-brand-yellow rounded-full animate-pulse"></div>
                    {t.latestNews}
                  </div>
                  <div className="flex animate-scroll hover:pause whitespace-nowrap pl-[160px] md:pl-[200px]">
                    <span className="mx-8 text-brand-green font-bold text-sm md:text-base tracking-tight">{dynamicTicker}</span>
                    <span className="mx-8 text-brand-green font-bold text-sm md:text-base tracking-tight">{dynamicTicker}</span>
                    <span className="mx-8 text-brand-green font-bold text-sm md:text-base tracking-tight">{dynamicTicker}</span>
                    <span className="mx-8 text-brand-green font-bold text-sm md:text-base tracking-tight">{dynamicTicker}</span>
                  </div>
                </div>

                {/* Grid Section */}
                <section className="max-w-7xl mx-auto px-5 md:px-8 mt-12 md:mt-20">
                  <div className="grid lg:grid-cols-12 gap-8 md:gap-12">
                    <div className="lg:col-span-4">
                      <div className="bento-card p-6 md:p-8 relative overflow-hidden h-full">
                        <div className="flex items-center justify-between mb-8 pb-6 border-b border-gray-100">
                          <div className="flex items-center gap-3">
                            <div className="bg-brand-gradient p-3 rounded-xl text-white shadow-lg shadow-brand-green/20">
                              <Pin size={18} className="rotate-45" />
                            </div>
                            <h2 className="text-xl md:text-2xl font-black text-slate-900 tracking-tight font-display">{t.noticeBoard}</h2>
                          </div>
                          {isAdmin && (
                            <button 
                               onClick={() => setIsAddingNotice(true)}
                               className="p-2.5 bg-brand-gradient text-white rounded-xl hover:opacity-90 transition-all shadow-lg shadow-brand-green/20"
                            >
                              <Plus size={18} />
                            </button>
                          )}
                        </div>

                        <AnimatePresence>
                          {isAddingNotice && (
                            <motion.div 
                              initial={{ height: 0, opacity: 0 }}
                              animate={{ height: 'auto', opacity: 1 }}
                              exit={{ height: 0, opacity: 0 }}
                              className="mb-8 p-5 bg-brand-accent rounded-3xl border border-brand-green/10 space-y-4 overflow-hidden"
                            >
                              <input 
                                type="text" 
                                placeholder="English text..."
                                className="w-full px-4 py-3 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-brand-green/20 outline-none transition-all"
                                value={newNotice.text_en}
                                onChange={(e) => setNewNotice({...newNotice, text_en: e.target.value})}
                              />
                              <input 
                                type="text" 
                                placeholder="Bengali text..."
                                className="w-full px-4 py-3 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-brand-green/20 outline-none transition-all"
                                value={newNotice.text_bn}
                                onChange={(e) => setNewNotice({...newNotice, text_bn: e.target.value})}
                              />
                              <div className="flex gap-2">
                                <button 
                                  onClick={addNotice}
                                  className="flex-grow bg-brand-gradient text-white py-3 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-brand-green/20"
                                >
                                  {lang === 'en' ? 'Publish' : 'প্রকাশ করুন'}
                                </button>
                                <button 
                                  onClick={() => setIsAddingNotice(false)}
                                  className="px-4 py-3 bg-white border border-slate-200 rounded-xl text-[10px] font-black uppercase tracking-widest text-slate-400"
                                >
                                  {lang === 'en' ? 'Cancel' : 'বাতিল'}
                                </button>
                              </div>
                            </motion.div>
                          )}
                        </AnimatePresence>

                        <div className="space-y-4">
                          {notices.slice(0, 4).map((n, i) => (
                            <a 
                              key={i} 
                              href={n.file || '#'} 
                              target={n.file ? "_blank" : "_self"}
                              rel="noopener noreferrer"
                              className={`flex gap-4 group/item p-4 rounded-2xl hover:bg-slate-50 border border-transparent hover:border-slate-100 transition-all ${n.file ? 'cursor-pointer' : 'cursor-default'}`}
                            >
                              <div className="bg-brand-accent text-brand-green text-[10px] font-black px-3 py-2 rounded-xl h-fit border border-brand-green/10 group-hover/item:bg-brand-gradient group-hover/item:text-white transition-all text-center leading-none uppercase shrink-0">
                                {n.date.split('-').slice(1).reverse().join('/')}
                              </div>
                              <div className="flex-grow">
                                <p className="text-sm font-bold text-slate-700 leading-snug group-hover/item:text-brand-green transition-colors flex items-center flex-wrap gap-2">
                                  {lang === 'en' ? n.text_en : n.text_bn}
                                  {n.file && <Download size={12} className="text-brand-yellow" />}
                                </p>
                              </div>
                            </a>
                          ))}
                        </div>

                        <div className="mt-8 pt-6 border-t border-slate-100">
                          <button 
                            onClick={() => setActiveTab('notices')}
                            className="w-full py-4 rounded-2xl bg-brand-accent text-brand-green font-black text-[10px] uppercase tracking-widest hover:bg-brand-gradient hover:text-white transition-all flex items-center justify-center gap-3 group/all shadow-sm"
                          >
                            {t.allNotices}
                            <ChevronRight size={14} className="group-hover/all:translate-x-1 transition-transform" />
                          </button>
                        </div>
                      </div>
                    </div>

                    <div className="lg:col-span-8 space-y-8 md:space-y-12">
                      <div className="bg-brand-accent p-8 md:p-14 rounded-[2.5rem] md:rounded-[4rem] border border-brand-green/5 relative overflow-hidden group">
                         <Leaf className="absolute -top-10 -right-10 text-brand-green/5 w-40 md:w-80 h-40 md:h-80 rotate-12 group-hover:rotate-45 transition-transform duration-1000" />
                         <h2 className="text-3xl md:text-5xl font-black text-slate-900 mb-6 tracking-tight font-display">{t.aboutTitle}</h2>
                         <p className="text-base md:text-xl text-slate-600 font-medium leading-relaxed italic border-l-4 border-brand-gradient pl-6 md:pl-10">
                           &quot;{t.aboutDesc}&quot;
                         </p>
                      </div>
                      <div className="grid sm:grid-cols-3 gap-6">
                        <FeatureItem icon={FlaskConical} label={t.modernLabs} />
                        <FeatureItem icon={Users} label={t.expertTeachers} />
                        <FeatureItem icon={Briefcase} label={t.careerSupport} />
                      </div>
                    </div>
                  </div>
                </section>
              </div>
            )}

            {activeTab === 'about' && (
              <section className="max-w-7xl mx-auto px-5 py-20">
                <div className="mobile-center text-center mb-16 max-w-3xl mx-auto space-y-4">
                   <h2 className="text-3xl md:text-7xl font-black bg-brand-gradient bg-clip-text text-transparent tracking-tighter uppercase leading-none font-display">
                     {t.aboutTitle}
                   </h2>
                   <div className="h-1.5 w-24 bg-brand-yellow mx-auto rounded-full shadow-lg md:mx-0"></div>
                </div>
                
                <div className="space-y-16">
                  {/* History */}
                  <motion.div 
                    initial={{ opacity: 0, y: 20 }}
                    whileInView={{ opacity: 1, y: 0 }}
                    viewport={{ once: true }}
                    className="bg-white p-10 md:p-16 rounded-[4rem] shadow-2xl shadow-brand-green/5 border border-brand-green/5 relative overflow-hidden"
                  >
                    <div className="relative z-10">
                      <h3 className="text-3xl font-black text-brand-green mb-8 flex items-center gap-4">
                        <Clock className="w-8 h-8 text-brand-yellow" />
                        {t.historyTitle}
                      </h3>
                      <div className="space-y-6">
                        {t.historyDesc.split('\n\n').map((para, idx) => (
                          <p key={idx} className="text-lg text-gray-600 font-medium leading-relaxed">
                            {para}
                          </p>
                        ))}
                      </div>
                    </div>
                  </motion.div>

              {/* Grid: Mission & Vision */}
              <div className="grid lg:grid-cols-2 gap-8">
                {/* Mission */}
                <motion.div 
                  initial={{ opacity: 0, x: -20 }}
                  whileInView={{ opacity: 1, x: 0 }}
                  viewport={{ once: true }}
                  className="bg-white p-8 md:p-12 rounded-[2.5rem] shadow-2xl shadow-brand-green/5 border border-brand-green/5"
                >
                  <h3 className="text-2xl font-black text-brand-green mb-6 flex items-center gap-3 font-display">
                    <Target className="w-6 h-6 text-brand-yellow" />
                    {t.missionTitle}
                  </h3>
                  <div className="space-y-4">
                    {t.missionDesc.split('\n\n').map((para, idx) => (
                      <p key={idx} className="text-base text-gray-600 font-medium leading-relaxed">
                        {para}
                      </p>
                    ))}
                  </div>
                </motion.div>

                {/* Vision */}
                <motion.div 
                  initial={{ opacity: 0, x: 20 }}
                  whileInView={{ opacity: 1, x: 0 }}
                  viewport={{ once: true }}
                  className="bg-brand-green text-white p-8 md:p-12 rounded-[2.5rem] shadow-2xl shadow-brand-green/20 relative overflow-hidden group"
                >
                  <Leaf className="absolute -top-10 -right-10 text-white/5 w-60 h-60 rotate-12 group-hover:scale-110 transition-transform duration-1000" />
                  <div className="relative z-10">
                    <h3 className="text-2xl font-black text-brand-yellow mb-6 flex items-center gap-3 font-display">
                      <Globe className="w-6 h-6" />
                      {t.visionTitle}
                    </h3>
                    <div className="space-y-4">
                      {t.visionDesc.split('\n\n').map((para, idx) => (
                        <p key={idx} className="text-base text-white/80 font-medium leading-relaxed">
                          {para}
                        </p>
                      ))}
                    </div>
                  </div>
                </motion.div>
              </div>
                </div>
              </section>
            )}

            {activeTab === 'students' && (
              <section className="max-w-7xl mx-auto px-5 py-20">
                <div className="mobile-center text-center mb-12 max-w-3xl mx-auto space-y-4">
                  <h2 className="text-3xl md:text-7xl font-black bg-brand-gradient bg-clip-text text-transparent tracking-tighter uppercase leading-none">
                    {t.registeredStudents}
                  </h2>
                  <div className="h-1.5 w-24 bg-brand-yellow rounded-full shadow-lg mx-auto md:mx-0"></div>
                  <p className="text-sm md:text-xl text-slate-500 font-medium">{t.studentDirMsg}</p>
                </div>

                {/* Filter Engine */}
                <div className="bg-white p-8 rounded-[3rem] shadow-2xl shadow-gray-100 border border-brand-green/10 mb-12 grid md:grid-cols-4 gap-6">
                   <div className="relative group">
                      <Search className="absolute left-5 top-1/2 -translate-y-1/2 text-gray-400 group-focus-within:text-brand-green transition-colors" size={20} />
                      <input 
                        type="text" 
                        placeholder={t.searchPlaceholder}
                        className="w-full pl-14 pr-6 py-5 bg-gray-50 rounded-2xl border-none focus:ring-2 focus:ring-brand-green/20 font-bold text-sm text-gray-700 placeholder:text-gray-400/80 outline-none transition-all"
                        value={searchQuery}
                        onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(1); }}
                      />
                   </div>
                   <EnhancedSelect 
                     value={semesterFilter} 
                     onChange={setSemesterFilter} 
                     options={["1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th"]}
                     placeholder={t.filterSemester}
                   />
                   <EnhancedSelect 
                     value={shiftFilter} 
                     onChange={setShiftFilter} 
                     options={["1st", "2nd"]}
                     placeholder={t.filterShift}
                   />
                   <EnhancedSelect 
                     value={groupFilter} 
                     onChange={setGroupFilter} 
                     options={["A", "B", "C"]}
                     placeholder={t.filterGroup}
                   />
                </div>

                {/* Records Table */}
                <div className="bg-white rounded-[3rem] shadow-2xl shadow-brand-green/5 border border-brand-green/5 overflow-hidden relative">
                   <div className="overflow-x-auto">
                     <table className="w-full text-left border-collapse min-w-[800px]">
                       <thead>
                         <tr className="bg-brand-green text-white font-black uppercase text-xs tracking-[0.2em]">
                           <th className="px-8 py-6">ID</th>
                           <th className="px-8 py-6">Identity</th>
                           <th className="px-8 py-6">Academic Path</th>
                           <th className="px-8 py-6">Communication</th>
                         </tr>
                       </thead>
                       <tbody className="divide-y divide-gray-100 font-bold">
                         {paginatedStudents.map((s, i) => (
                           <tr key={i} className="hover:bg-brand-green/5 transition-all group">
                             <td className="px-8 py-6 text-gray-300 font-mono text-sm leading-none tabular-nums group-hover:text-brand-green transition-colors">
                               #{(currentPage-1)*itemsPerPage + i + 1}
                             </td>
                             <td className="px-8 py-6">
                                <div className="flex items-center gap-5">
                                  <div className="relative">
                                    <div className="w-16 h-16 rounded-[1.5rem] bg-gray-100 overflow-hidden ring-4 ring-gray-50 group-hover:ring-brand-green/10 transition-all">
                                      <img src={s.image || `https://ui-avatars.com/api/?name=${encodeURIComponent(s.name)}&background=116530&color=fff`} className="w-full h-full object-cover" alt="Avatar" />
                                    </div>
                                    <div className="absolute -bottom-1 -right-1 w-4 h-4 bg-brand-green border-2 border-white rounded-full"></div>
                                  </div>
                                  <div>
                                    <h4 className="text-lg font-black text-gray-900 leading-none group-hover:text-brand-green transition-colors">{s.name}</h4>
                                    <div className="flex items-center gap-2 mt-2">
                                      <span className="text-[10px] bg-gray-100 px-2 py-0.5 rounded-md text-gray-500 font-black uppercase">{s.roll}</span>
                                      <span className="text-[10px] text-brand-green opacity-60">REG: {s.registration}</span>
                                    </div>
                                  </div>
                                </div>
                             </td>
                             <td className="px-8 py-6">
                                <div className="flex items-center gap-2">
                                   <div className="bg-white border border-gray-100 shadow-sm p-3 rounded-2xl text-center min-w-[70px]">
                                      <p className="text-[8px] text-gray-300 uppercase font-black mb-0.5 tracking-tighter">Semester</p>
                                      <p className="text-sm font-black text-brand-green">{s.semester}</p>
                                   </div>
                                   <div className="bg-white border border-gray-100 shadow-sm p-3 rounded-2xl text-center min-w-[70px]">
                                      <p className="text-[8px] text-gray-300 uppercase font-black mb-0.5 tracking-tighter">Shift</p>
                                      <p className="text-sm font-black text-blue-700">{s.shift}</p>
                                   </div>
                                   <div className="bg-white border border-gray-100 shadow-sm p-3 rounded-2xl text-center min-w-[70px]">
                                      <p className="text-[8px] text-gray-300 uppercase font-black mb-0.5 tracking-tighter">Group</p>
                                      <p className="text-sm font-black text-purple-700">{s.group}</p>
                                   </div>
                                   {s.session && (
                                     <div className="bg-white border border-gray-100 shadow-sm p-3 rounded-2xl text-center min-w-[70px]">
                                        <p className="text-[8px] text-gray-300 uppercase font-black mb-0.5 tracking-tighter">{t.session}</p>
                                        <p className="text-sm font-black text-orange-600">{s.session}</p>
                                     </div>
                                   )}
                                </div>
                             </td>
                             <td className="px-8 py-6">
                                <div className="space-y-2">
                                  <a href={`tel:${s.mobile}`} className="flex items-center gap-3 text-sm text-gray-700 hover:text-brand-green transition-all whitespace-nowrap">
                                    <div className="p-2 bg-brand-green/10 text-brand-green rounded-lg">
                                      <Phone size={14} />
                                    </div>
                                    {s.mobile}
                                  </a>
                                  <div className="flex items-center gap-3 text-xs text-gray-400 whitespace-nowrap">
                                    <div className="p-2 bg-gray-50 text-gray-400 rounded-lg">
                                      <Users size={14} />
                                    </div>
                                    <span className="font-medium tracking-tight">Guardian: {s.guardian}</span>
                                  </div>
                                </div>
                             </td>
                           </tr>
                         ))}
                       </tbody>
                     </table>
                   </div>
                   
                   {/* Table Footer */}
                   <div className="p-8 bg-gray-50/50 border-t border-gray-100 flex flex-col sm:flex-row justify-between items-center gap-6">
                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-[0.3em]">
                        Showing page {currentPage} of {totalPages || 1}
                      </p>
                      {totalPages > 1 && (
                        <div className="flex items-center gap-2">
                          <button 
                            disabled={currentPage === 1}
                            onClick={() => setCurrentPage(c => Math.max(1, c - 1))}
                            className="bg-white border border-gray-200 p-3 rounded-xl disabled:opacity-30 hover:border-brand-green hover:text-brand-green transition-all shadow-sm"
                          >
                            <ChevronRight size={18} className="rotate-180" />
                          </button>
                          {Array.from({length: totalPages}).map((_, i) => {
                            if (Math.abs(currentPage - (i + 1)) <= 1 || i + 1 === 1 || i + 1 === totalPages) {
                              return (
                                <button 
                                  key={i}
                                  onClick={() => setCurrentPage(i + 1)}
                                  className={`w-12 h-12 rounded-xl font-black text-xs transition-all ${currentPage === i + 1 ? 'bg-brand-green text-white shadow-xl shadow-brand-green/30' : 'bg-white text-gray-400 border border-gray-200 hover:border-brand-green hover:text-brand-green shadow-sm'}`}
                                >
                                  {i + 1}
                                </button>
                              );
                            }
                            return null;
                          })}
                          <button 
                            disabled={currentPage === totalPages}
                            onClick={() => setCurrentPage(c => Math.min(totalPages, c + 1))}
                            className="bg-white border border-gray-200 p-3 rounded-xl disabled:opacity-30 hover:border-brand-green hover:text-brand-green transition-all shadow-sm"
                          >
                            <ChevronRight size={18} />
                          </button>
                        </div>
                      )}
                   </div>
                </div>

                {/* Additional Info Grid from User Request */}
                <div className="grid md:grid-cols-2 gap-10 mt-16">
                  <div className="bg-white p-10 rounded-[3rem] shadow-2xl shadow-brand-green/5 border border-brand-green/5">
                    <div className="flex items-center gap-4 mb-10">
                      <div className="bg-brand-green/10 p-4 rounded-2xl text-brand-green">
                        <Users size={24} />
                      </div>
                      <h3 className="text-2xl font-black text-gray-900 tracking-tighter uppercase">
                        {lang === 'en' ? 'Current Batches' : 'বর্তমান ব্যাচসমূহ'}
                      </h3>
                    </div>
                    <div className="space-y-6">
                      {[
                        { label: '2nd Semester', session: '2025-2026', labelBn: '২য় সেমিস্টার', sessionBn: '২০২৫-২০২৬', tag: '1st' },
                        { label: '3rd Semester', session: '2024-2025', labelBn: '৩য় সেমিস্টার', sessionBn: '২০২৪-২০২৫', tag: '3rd' },
                        { label: '5th Semester', session: '2023-2024', labelBn: '৫ম সেমিস্টার', sessionBn: '২০২৩-২০২৪', tag: '5th' },
                        { label: '7th Semester', session: '2022-2023', labelBn: '৭ম সেমিস্টার', sessionBn: '২০২২-২০২৩', tag: '7th' }
                      ].map((batch, idx) => (
                        <div key={idx} className="flex items-center gap-5 p-4 bg-gray-50 rounded-2xl border border-gray-100 hover:border-brand-green/20 transition-all">
                          <div className="w-12 h-12 bg-brand-green text-white rounded-xl flex items-center justify-center font-black text-xs">
                            {batch.tag}
                          </div>
                          <div>
                            <p className="font-bold text-gray-800">{lang === 'en' ? batch.label : batch.labelBn}</p>
                            <p className="text-[10px] text-gray-400 font-black uppercase tracking-wider">Session: {lang === 'en' ? batch.session : batch.sessionBn}</p>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  <div className="bg-brand-green text-white p-10 rounded-[3rem] shadow-2xl shadow-brand-green/20 relative overflow-hidden group">
                    <Leaf className="absolute -top-10 -right-10 text-white/5 w-60 h-60 rotate-12 group-hover:scale-110 transition-transform duration-1000" />
                    <div className="relative z-10">
                      <div className="flex items-center gap-4 mb-8">
                        <div className="bg-white/20 p-4 rounded-2xl backdrop-blur-xl">
                          <Globe size={24} className="text-brand-yellow" />
                        </div>
                        <h3 className="text-2xl font-black tracking-tighter uppercase">
                          {lang === 'en' ? 'Club & Activities' : 'ক্লাব ও কার্যক্রম'}
                        </h3>
                      </div>
                      <p className="text-xl text-white/80 font-medium leading-relaxed italic border-l-4 border-brand-yellow pl-8 mb-10">
                        {lang === 'en' 
                          ? "Students regularly participate in Food Fairs, Skills Competitions, and Industrial Tours to gain real-world exposure." 
                          : "শিক্ষার্থীরা বাস্তব অভিজ্ঞতা অর্জনের জন্য নিয়মিত খাদ্য মেলা, দক্ষতা প্রতিযোগিতা এবং শিল্প কারখানায় ভ্রমণে অংশগ্রহণ করে।"}
                      </p>
                      <button 
                         onClick={() => setActiveTab('home')}
                         className="bg-brand-yellow text-brand-green px-10 py-4 rounded-2xl font-black text-xs uppercase tracking-widest hover:scale-105 transition-all shadow-xl shadow-black/10"
                      >
                         {lang === 'en' ? 'Learn More' : 'আরও জানুন'}
                      </button>
                    </div>
                  </div>
                </div>
              </section>
            )}

            {(activeTab === 'teachers' || activeTab === 'staff') && (
              <section className="max-w-7xl mx-auto px-5 py-20">
                <div className="mobile-center text-center mb-12 max-w-3xl mx-auto space-y-4">
                   <h2 className="text-3xl md:text-7xl font-black bg-brand-gradient bg-clip-text text-transparent tracking-tighter uppercase leading-none">
                     {activeTab === 'teachers' ? (lang === 'en' ? 'Faculty Members' : 'ফ্যাকাল্টি মেম্বারগণ') : (lang === 'en' ? 'Office & Lab Staff' : 'অফিস ও ল্যাব স্টাফ')}
                   </h2>
                   <div className="h-1.5 w-24 bg-brand-yellow rounded-full shadow-lg mx-auto md:mx-0"></div>
                   <p className="text-base md:text-2xl text-slate-500 font-medium italic border-l-4 border-brand-yellow pl-6 text-left">
                     {activeTab === 'teachers' ? t.facultyMsg : t.staffMsg}
                   </p>
                </div>

                {/* Search & Filter UI */}
                <div className="max-w-4xl mx-auto mb-16 flex flex-col md:flex-row gap-4 bg-white p-6 rounded-3xl shadow-xl shadow-brand-green/5 border border-brand-green/10">
                  <div className="flex-1 relative">
                    <Search className="absolute left-6 top-1/2 -translate-y-1/2 w-5 h-5 text-gray-400 font-bold" />
                    <input 
                      type="text"
                      placeholder={lang === 'en' ? "Search by name or position..." : "নাম বা পদবী দিয়ে খুঁজুন..."}
                      className="w-full pl-14 pr-4 py-4 bg-gray-50 border border-gray-200 rounded-2xl focus:ring-2 focus:ring-brand-green focus:border-transparent transition-all outline-none font-medium"
                      value={teacherSearch}
                      onChange={(e) => setTeacherSearch(e.target.value)}
                    />
                  </div>
                  <select 
                    className="px-8 py-4 bg-gray-50 border border-gray-200 rounded-2xl focus:ring-2 focus:ring-brand-green focus:border-transparent transition-all outline-none cursor-pointer font-bold text-gray-700"
                    value={teacherFilter}
                    onChange={(e) => setTeacherFilter(e.target.value)}
                  >
                    <option value="All">{lang === 'en' ? "All Categories" : "সকল বিভাগ"}</option>
                    {activeTab === 'teachers' ? (
                      <>
                        <option value="Head">{lang === 'en' ? "Dept. Head" : "বিভাগীয় প্রধান"}</option>
                        <option value="Instructor">{lang === 'en' ? "Instructor" : "ইনস্ট্রাক্টর"}</option>
                        <option value="Junior">{lang === 'en' ? "Junior Instructor" : "জুনিয়র ইনস্ট্রাক্টর"}</option>
                      </>
                    ) : (
                      <>
                        <option value="Craft">{lang === 'en' ? "Craft Instructor" : "ক্রাফট ইন্সট্রাক্টর"}</option>
                        <option value="Assistant">{lang === 'en' ? "Office Assistant" : "অফিস সহকারী"}</option>
                      </>
                    )}
                  </select>
                </div>

                {filteredTeachers.length > 0 ? (
                  <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-6">
                    {filteredTeachers.map((person, i) => (
                      <motion.div 
                        key={i}
                        initial={{ opacity: 0, y: 20 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ delay: i * 0.05 }}
                        className="group bg-white rounded-[2rem] p-6 shadow-xl border border-gray-100"
                      >
                        <div className="absolute top-0 right-0 w-24 h-24 bg-brand-green/5 rounded-bl-[3rem] -mr-12 -mt-12 group-hover:bg-brand-green/10 transition-colors"></div>
                        
                        <div className="relative mb-6">
                          <div className="w-32 h-32 rounded-full border-2 border-brand-green/10 p-1 group-hover:border-brand-green transition-all duration-500 overflow-hidden">
                            <div className="w-full h-full rounded-full overflow-hidden ring-2 ring-white shadow-inner">
                              <img 
                                src={person.image || `https://ui-avatars.com/api/?name=${encodeURIComponent(person.name_en)}&background=116530&color=fff`} 
                                alt={person.name_en}
                                className="w-full h-full object-cover transition-all duration-1000 group-hover:scale-110"
                                referrerPolicy="no-referrer"
                              />
                            </div>
                          </div>
                          <div className="absolute -bottom-1 right-2 bg-brand-yellow text-brand-green p-2 rounded-xl shadow-lg border-2 border-white group-hover:rotate-12 transition-transform">
                            <Briefcase size={14} />
                          </div>
                        </div>

                        <div className="space-y-3 w-full relative z-10">
                          <h3 className="text-lg font-black text-gray-900 leading-tight group-hover:text-brand-green transition-colors font-display">
                            {lang === 'en' ? person.name_en : person.name_bn}
                          </h3>
                          <div className="inline-block px-3 py-1 bg-brand-green/5 text-brand-green text-[9px] font-black uppercase tracking-widest rounded-full">
                            {lang === 'en' ? person.position_en : person.position_bn}
                          </div>
                          
                          <div className="pt-4 space-y-2 border-t border-gray-100">
                            <div className="flex items-center justify-center gap-2 text-slate-500 text-[10px] font-bold bg-gray-50/50 py-1.5 rounded-lg">
                              <Phone size={12} className="text-brand-green" />
                              <span className="tracking-tighter">{person.mobile}</span>
                            </div>
                            <div className="flex items-center justify-center gap-2 text-slate-500 text-[10px] font-bold bg-gray-50/50 py-1.5 rounded-lg">
                              <Mail size={12} className="text-brand-green" />
                              <span className="truncate max-w-[150px] tracking-tighter">{person.email}</span>
                            </div>
                          </div>

                          <div className="flex justify-center gap-4 pt-4">
                            <a 
                              href={`tel:${person.mobile}`} 
                              className="p-4 bg-brand-green text-white rounded-2xl transition-all hover:scale-110 active:scale-95 shadow-xl shadow-brand-green/20 hover:bg-brand-green/90"
                              title="Call Now"
                            >
                              <Phone size={24} />
                            </a>
                            <a 
                              href={`mailto:${person.email}`} 
                              className="p-4 bg-brand-yellow text-brand-green rounded-2xl transition-all hover:scale-110 active:scale-95 shadow-xl shadow-brand-yellow/20 hover:bg-brand-yellow/80"
                              title="Send Email"
                            >
                              <Mail size={24} />
                            </a>
                          </div>
                        </div>
                      </motion.div>
                    ))}
                  </div>
                ) : (
                  <div className="text-center py-32 bg-gray-50/50 rounded-[4rem] border-4 border-dashed border-gray-200">
                    <Search className="w-20 h-20 text-gray-200 mx-auto mb-6" />
                    <h3 className="text-3xl font-black text-gray-800 mb-2 uppercase tracking-tighter">
                       {lang === 'en' ? "Empty Lab Results" : "কোনো ফলাফল পাওয়া যায়নি"}
                    </h3>
                    <p className="text-gray-400 font-bold max-w-sm mx-auto uppercase text-xs tracking-widest leading-relaxed">
                      {lang === 'en' ? "Modify your parameters to locate the registered faculty or personnel." : "নিবন্ধিত অনুষদ বা কর্মীদের খুঁজে পেতে আপনার প্যারামিটার পরিবর্তন করুন।"}
                    </p>
                  </div>
                )}
              </section>
            )}

            {activeTab === 'labs' && (
              <section className="max-w-7xl mx-auto px-4 py-20 divide-y-2 divide-brand-green/5">
                 <h2 className="text-6xl font-black mb-24 text-center text-gray-900 tracking-tighter uppercase">{t.labs}</h2>
                 {labs.map((lab, i) => (
                   <div key={i} className={`flex flex-col lg:flex-row gap-16 py-24 group ${i % 2 !== 0 ? 'lg:flex-row-reverse' : ''}`}>
                      <div className="lg:w-1/2 grid grid-cols-2 gap-4 relative">
                        <div className="absolute -inset-10 bg-brand-green/5 rounded-full blur-3xl pointer-events-none group-hover:bg-brand-green/10 transition-colors"></div>
                        {lab.images.map((img, idx) => (
                          <div key={idx} className={`rounded-[2.5rem] overflow-hidden shadow-2xl relative z-10 ${idx === 0 ? 'translate-y-6' : '-translate-y-6'}`}>
                            <img src={img || 'https://images.unsplash.com/photo-1514432324607-a09d9b4aefdd?auto=format&fit=crop&q=80'} className="w-full h-full object-cover aspect-[4/3] group-hover:scale-110 transition-transform duration-1000" alt="Lab" />
                          </div>
                        ))}
                      </div>
                      <div className="lg:w-1/2 flex flex-col justify-center">
                        <div className="inline-flex items-center gap-3 bg-brand-green/5 text-brand-green px-6 py-2 rounded-full font-black text-xs uppercase tracking-widest mb-8 border border-brand-green/10">
                           <FlaskConical size={18} /> Research Facility 0{i+1}
                        </div>
                        <h3 className="text-4xl md:text-6xl font-black text-gray-900 tracking-tighter mb-8 leading-none">
                          {lang === 'en' ? lab.name_en : lab.name_bn}
                        </h3>
                        <p className="text-xl text-gray-400 font-medium leading-relaxed italic border-l-8 border-brand-yellow pl-8 mb-12">
                          {lang === 'en' ? lab.desc_en : lab.desc_bn}
                        </p>
                        <ul className="grid sm:grid-cols-2 gap-4">
                           <LabPoint text="Precision Equipment" />
                           <LabPoint text="Departmental Oversight" />
                           <LabPoint text="International Standards" />
                           <LabPoint text="Technical Workshops" />
                        </ul>
                      </div>
                   </div>
                 ))}
              </section>
            )}

            {activeTab === 'syllabus' && (
              <section className="max-w-7xl mx-auto px-5 py-20">
                <div className="mobile-center text-center mb-16 max-w-3xl mx-auto space-y-4">
                  <h2 className="text-3xl md:text-7xl font-black bg-brand-gradient bg-clip-text text-transparent tracking-tighter uppercase leading-none">
                    {t.syllabus}
                  </h2>
                  <div className="h-1.5 w-24 bg-brand-yellow rounded-full shadow-lg mx-auto md:mx-0"></div>
                  <p className="text-sm md:text-xl text-slate-500 font-medium">{t.syllabusDesc}</p>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-8">
                  {[1, 2, 3, 4, 5, 6, 7, 8].map((sem) => {
                    const isLast = sem === 8;
                    const semKey = sem.toString();
                    const subjects = semesters[semKey] || [];
                    
                    return (
                      <motion.div
                        key={sem}
                        initial={{ opacity: 0, y: 20 }}
                        whileInView={{ opacity: 1, y: 0 }}
                        viewport={{ once: true }}
                        transition={{ delay: sem * 0.05 }}
                        className="group relative"
                      >
                        <div className="h-full bg-white rounded-[3rem] p-8 border border-gray-100 shadow-xl shadow-brand-green/[0.03] transition-all duration-500 hover:shadow-2xl hover:shadow-brand-green/20 hover:-translate-y-2 flex flex-col items-center text-center overflow-hidden">
                          {/* Background Glow */}
                          <div className={`absolute inset-0 bg-gradient-to-br transition-opacity duration-500 opacity-0 group-hover:opacity-100 ${isLast ? 'from-orange-50 to-transparent' : 'from-brand-green/5 to-transparent'}`} />
                          
                          <div className={`relative z-10 w-20 h-20 rounded-[2.5rem] flex items-center justify-center mb-6 shadow-xl transition-transform duration-500 group-hover:scale-110 group-hover:rotate-6 ${isLast ? 'bg-orange-500 text-white shadow-orange-500/30' : 'bg-brand-green text-white shadow-brand-green/30'}`}>
                            {isLast ? <Briefcase className="w-10 h-10" /> : <FileText className="w-10 h-10" />}
                          </div>

                          <h3 className="relative z-10 text-2xl font-black text-gray-900 tracking-tighter uppercase leading-tight mb-2">
                            {lang === 'en' ? `${sem}${sem === 1 ? 'st' : sem === 2 ? 'nd' : sem === 3 ? 'rd' : 'th'} Semester` : `${sem === 1 ? '১ম' : sem === 2 ? '২য়' : sem === 3 ? '৩য়' : sem === 4 ? '৪র্থ' : sem === 5 ? '৫ম' : sem === 6 ? '৬ষ্ঠ' : sem === 7 ? '৭ম' : '৮ম'} সেমিস্টার`}
                          </h3>
                          
                          <div className="relative z-10 mt-auto pt-6 border-t border-gray-100 w-full">
                            {subjects.length > 0 ? (
                              <div className="space-y-3">
                                {subjects.map((sub, idx) => (
                                  <button 
                                    key={idx}
                                    onClick={() => setSelectedCourse({ ...sub, semester: sem })}
                                    className="w-full flex items-center justify-between p-3 rounded-2xl bg-gray-50 hover:bg-brand-green hover:text-white transition-all text-xs font-black uppercase tracking-widest text-brand-green group/btn"
                                  >
                                    <span className="truncate pr-2">{sub.name}</span>
                                    <ChevronRight size={14} className="shrink-0 group-hover/btn:translate-x-1 transition-transform" />
                                  </button>
                                ))}
                              </div>
                            ) : (
                              <button className="w-full py-4 rounded-2xl bg-gray-50 text-gray-400 text-[10px] font-black uppercase tracking-[0.2em] pointer-events-none">
                                Comming Soon
                              </button>
                            )}
                          </div>
                        </div>
                      </motion.div>
                    );
                  })}
                </div>
              </section>
            )}

            {activeTab === 'notices' && (
              <section className="max-w-4xl mx-auto px-4 py-20">
                <div className="text-center mb-16">
                  <h2 className="text-6xl font-black text-gray-900 tracking-tighter uppercase mb-6">{t.allNotices}</h2>
                  <div className="h-2 w-32 bg-brand-yellow mx-auto rounded-full shadow-lg"></div>
                </div>

                <div className="bg-white p-6 rounded-[2.5rem] shadow-xl border border-brand-green/5 mb-10 flex items-center gap-4">
                  <Search className="text-gray-400 ml-4" size={24} />
                  <input 
                    type="text" 
                    placeholder={t.noticeSearch}
                    className="flex-grow bg-transparent border-none focus:ring-0 font-bold text-lg text-gray-700 outline-none placeholder:text-gray-300"
                    value={noticeSearch}
                    onChange={(e) => setNoticeSearch(e.target.value)}
                  />
                </div>

                <div className="bg-white rounded-[3rem] shadow-2xl shadow-brand-green/5 border border-brand-green/5 overflow-hidden">
                  <div className="divide-y divide-gray-50">
                    {filteredAllNotices.length > 0 ? (
                      filteredAllNotices.map((n, i) => (
                        <motion.a 
                          key={i} 
                          initial={{ opacity: 0, x: -20 }}
                          animate={{ opacity: 1, x: 0 }}
                          transition={{ delay: i * 0.05 }}
                          href={n.file || '#'}
                          target={n.file ? "_blank" : "_self"}
                          rel="noopener noreferrer"
                          className="flex items-center gap-8 p-10 hover:bg-brand-green/5 transition-all group"
                        >
                          <div className="flex flex-col items-center justify-center bg-gray-50 text-brand-green w-24 h-24 rounded-[2rem] border border-gray-100 group-hover:bg-brand-green group-hover:text-white transition-all shadow-sm shrink-0">
                            <span className="text-2xl font-black">{n.date.split('-')[2]}</span>
                            <span className="text-[10px] font-black uppercase tracking-widest">{new Date(n.date).toLocaleString('default', { month: 'short' })}</span>
                            <span className="text-[10px] opacity-60">{n.date.split('-')[0]}</span>
                          </div>
                          <div className="flex-grow space-y-3">
                            <p className="text-xl font-bold text-gray-800 leading-tight group-hover:text-brand-green transition-all">
                              {lang === 'en' ? n.text_en : n.text_bn}
                            </p>
                            <div className="flex items-center gap-4">
                              <span className="text-[10px] font-black bg-emerald-50 text-emerald-600 px-3 py-1 rounded-full uppercase tracking-widest">
                                Official Update
                              </span>
                              {n.file && (
                                <span className="flex items-center gap-2 text-[10px] font-black text-brand-yellow uppercase tracking-widest">
                                  <Download size={12} /> Attachment Available
                                </span>
                              )}
                            </div>
                          </div>
                          <div className="p-4 bg-gray-50 rounded-2xl text-gray-300 group-hover:bg-brand-green group-hover:text-white transition-all">
                            <ChevronRight size={24} />
                          </div>
                        </motion.a>
                      ))
                    ) : (
                      <div className="p-20 text-center">
                        <Search size={48} className="mx-auto text-gray-200 mb-6" />
                        <h3 className="text-2xl font-black text-gray-400 uppercase tracking-tighter">No Notices Found</h3>
                      </div>
                    )}
                  </div>
                </div>
              </section>
            )}

            {activeTab === 'student-portal' && (
              <section className="max-w-7xl mx-auto px-4 py-20">
                {!user ? (
                  <div className="max-w-md mx-auto bg-white rounded-[3rem] shadow-2xl overflow-hidden border border-gray-100">
                    <div className="bg-brand-green p-10 text-white text-center">
                      <div className="w-20 h-20 bg-white/10 rounded-[2rem] flex items-center justify-center mx-auto mb-6 backdrop-blur-xl">
                        <Users size={40} />
                      </div>
                      <h2 className="text-3xl font-black tracking-tighter mb-2">{authMode === 'login' ? t.login : t.register}</h2>
                      <p className="text-white/60 font-medium text-sm italic">{t.studentPortalTitle}</p>
                    </div>
                    
                    <form onSubmit={handleAuth} className="p-10 space-y-6">
                      {authMode === 'register' && (
                        <>
                          <div className="space-y-2">
                             <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Full Name</label>
                             <input name="name" required className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                          </div>
                          <div className="grid grid-cols-2 gap-4">
                             <div className="space-y-2">
                                <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Roll</label>
                                <input name="roll" required className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                             </div>
                             <div className="space-y-2">
                                <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Registration</label>
                                <input name="registration" required className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                             </div>
                          </div>
                          <div className="space-y-2">
                             <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Mobile</label>
                             <input name="mobile" required className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                          </div>
                          <div className="grid grid-cols-3 gap-4">
                             <div className="space-y-2">
                                <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Semester</label>
                                <select name="semester" required className="w-full px-4 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold text-xs appearance-none">
                                  {['1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th'].map(s => <option key={s} value={s}>{s}</option>)}
                                </select>
                             </div>
                             <div className="space-y-2">
                                <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Shift</label>
                                <select name="shift" required className="w-full px-4 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold text-xs appearance-none">
                                  {['1st', '2nd'].map(s => <option key={s} value={s}>{s}</option>)}
                                </select>
                             </div>
                             <div className="space-y-2">
                                <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Group</label>
                                <select name="group" required className="w-full px-4 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold text-xs appearance-none">
                                  {['A', 'B'].map(s => <option key={s} value={s}>{s}</option>)}
                                </select>
                             </div>
                             <div className="space-y-2">
                                <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Session</label>
                                <input name="session" placeholder="e.g. 2023-24" required className="w-full px-4 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold text-xs" />
                             </div>
                          </div>
                          <div className="space-y-2">
                             <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Profile Picture URL</label>
                             <input name="image" placeholder="https://..." className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                          </div>
                        </>
                      )}
                      <div className="space-y-2">
                         <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Email Address</label>
                         <input type="email" name="email" required className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                      </div>
                      <div className="space-y-2">
                         <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Password</label>
                         <input type="password" name="password" required className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                      </div>

                      {authError && <p className="text-red-500 text-xs font-bold bg-red-50 p-4 rounded-xl">{authError}</p>}

                      <button 
                        type="submit" 
                        disabled={authLoading}
                        className="w-full py-5 bg-brand-green text-white rounded-2xl font-black text-sm uppercase tracking-widest shadow-xl shadow-brand-green/20 hover:scale-[1.02] active:scale-95 transition-all flex items-center justify-center gap-3"
                      >
                        {authLoading ? '...' : (authMode === 'login' ? t.login : t.register)}
                        <ChevronRight size={18} />
                      </button>

                      <div className="text-center pt-4">
                        <button 
                          type="button"
                          onClick={() => setAuthMode(authMode === 'login' ? 'register' : 'login')}
                          className="text-gray-400 font-bold text-sm hover:text-brand-green transition-colors"
                        >
                          {authMode === 'login' ? "Don't have an account? Register" : "Already have an account? Login"}
                        </button>
                      </div>
                    </form>
                  </div>
                ) : (
                  <div className="space-y-12">
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-8 bg-white p-10 rounded-[3rem] shadow-xl border border-gray-100">
                       <div className="flex items-center gap-8">
                          <div className="relative group">
                            <img src={studentProfile?.basic.image || user.photoURL || `https://ui-avatars.com/api/?name=${encodeURIComponent(user.displayName || 'Student')}&background=116530&color=fff&size=128`} className="w-32 h-32 rounded-[2.5rem] border-4 border-brand-green/10 p-2 group-hover:border-brand-green transition-all" />
                            <div className="absolute -bottom-2 -right-2 bg-brand-yellow p-3 rounded-2xl border-4 border-white shadow-lg">
                              <Leaf size={20} className="text-brand-green" />
                            </div>
                          </div>
                          <div>
                            <div className="flex items-center gap-3 mb-2">
                               <h2 className="text-4xl font-black text-gray-900 tracking-tighter">{user.displayName}</h2>
                               {!user.emailVerified && <span className="bg-red-50 text-red-500 text-[10px] font-black px-2 py-1 rounded-md uppercase tracking-wider">Unverified</span>}
                            </div>
                            <p className="text-slate-400 font-bold text-lg">{user.email}</p>
                            <div className="flex gap-4 mt-6">
                               <button 
                                 onClick={() => printCV(studentProfile)} 
                                 className="px-8 py-3 bg-brand-green text-white rounded-xl font-black text-xs uppercase tracking-widest shadow-xl shadow-brand-green/10 flex items-center gap-2 hover:scale-105 transition-all"
                               >
                                 <Printer size={16} /> {t.printCV}
                               </button>
                               <button onClick={logout} className="px-8 py-3 bg-gray-50 text-gray-400 rounded-xl font-black text-xs uppercase tracking-widest border border-gray-100 hover:bg-red-50 hover:text-red-500 hover:border-red-100 transition-all flex items-center gap-2">
                                 <LogOut size={16} /> {t.logout}
                               </button>
                            </div>
                          </div>
                       </div>
                       
                       <div className="hidden lg:block">
                          <div className="bg-brand-green/5 p-8 rounded-[2rem] border border-brand-green/10 text-center">
                             <div className="text-3xl font-black text-brand-green mb-1">FT-2026</div>
                             <div className="text-[10px] font-black text-brand-green/60 uppercase tracking-widest">Portal Access Key</div>
                          </div>
                       </div>
                    </div>

                    {studentProfile && <ProfileEditor profile={studentProfile} onSave={saveProfile} lang={lang} t={t} />}
                  </div>
                )}
              </section>
            )}

            {activeTab === 'contact' && (
              <section className="max-w-7xl mx-auto px-5 py-20">
                <div className="mobile-center text-center mb-16 max-w-3xl mx-auto space-y-4">
                   <h2 className="text-3xl md:text-7xl font-black bg-brand-gradient bg-clip-text text-transparent tracking-tighter uppercase leading-none">
                     {t.contact}
                   </h2>
                   <div className="h-1.5 w-24 bg-brand-yellow rounded-full shadow-lg mx-auto md:mx-0"></div>
                   <p className="text-sm md:text-xl text-slate-500 font-medium italic">We're here to help you. Reach out for any departmental inquiries.</p>
                </div>

                <div className="grid md:grid-cols-2 gap-8 md:gap-12 items-start">
                   <div className="space-y-8">
                      <div className="bento-card p-8 space-y-8">
                         <div className="flex gap-6 items-start">
                            <div className="p-4 bg-brand-accent text-brand-green rounded-2xl shadow-sm">
                               <MapPin size={24} />
                            </div>
                            <div>
                               <h4 className="text-sm font-black text-slate-400 uppercase tracking-widest mb-2">Campus Location</h4>
                               <p className="text-lg font-bold text-slate-700">{t.location}</p>
                            </div>
                         </div>
                         <div className="flex gap-6 items-start">
                            <div className="p-4 bg-brand-accent text-brand-green rounded-2xl shadow-sm">
                               <Phone size={24} />
                            </div>
                            <div>
                               <h4 className="text-sm font-black text-slate-400 uppercase tracking-widest mb-2">Direct Hotline</h4>
                               <p className="text-lg font-bold text-slate-700">+880 1737256030</p>
                            </div>
                         </div>
                         <div className="flex gap-6 items-start">
                            <div className="p-4 bg-brand-accent text-brand-green rounded-2xl shadow-sm">
                               <Mail size={24} />
                            </div>
                            <div>
                               <h4 className="text-sm font-black text-slate-400 uppercase tracking-widest mb-2">Email Support</h4>
                               <p className="text-lg font-bold text-slate-700">food.chapaipoly@gmail.com</p>
                            </div>
                         </div>
                      </div>

                      <div className="bento-card p-8 bg-brand-gradient text-white relative overflow-hidden">
                         <div className="relative z-10">
                            <h3 className="text-2xl font-black mb-4 tracking-tighter">Emergency Help?</h3>
                            <p className="text-white/80 font-medium mb-6">Our administrative office is open Saturday-Thursday from 9:00 AM to 5:00 PM.</p>
                            <a href="tel:+8801737256030" className="inline-flex items-center gap-3 bg-brand-yellow text-brand-green px-8 py-4 rounded-xl font-black text-xs uppercase tracking-widest shadow-xl shadow-black/10 active:scale-95 transition-all">
                               Call Now <ChevronRight size={16} />
                            </a>
                         </div>
                         <Leaf className="absolute -bottom-10 -right-10 text-white/10 w-40 h-40" />
                      </div>
                   </div>

                   <div className="bento-card p-8 md:p-10">
                      <h3 className="text-2xl font-black text-slate-900 mb-8 flex items-center gap-3">
                         <Mail className="text-brand-green" /> {lang === 'en' ? 'Send Message' : 'বার্তা পাঠান'}
                      </h3>
                      <form className="space-y-6" onSubmit={(e) => e.preventDefault()}>
                         <div className="grid sm:grid-cols-2 gap-6">
                            <div className="space-y-2">
                               <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 ml-4">Full Name</label>
                               <input type="text" className="w-full px-6 py-4 bg-slate-50 rounded-2xl border border-slate-100 focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                            </div>
                            <div className="space-y-2">
                               <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 ml-4">Email</label>
                               <input type="email" className="w-full px-6 py-4 bg-slate-50 rounded-2xl border border-slate-100 focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                            </div>
                         </div>
                         <div className="space-y-2">
                            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 ml-4">Subject</label>
                            <input type="text" className="w-full px-6 py-4 bg-slate-50 rounded-2xl border border-slate-100 focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" />
                         </div>
                         <div className="space-y-2">
                            <label className="text-[10px] font-black uppercase tracking-widest text-slate-400 ml-4">Message</label>
                            <textarea rows={5} className="w-full px-6 py-4 bg-slate-50 rounded-2xl border border-slate-100 focus:ring-2 focus:ring-brand-green/20 outline-none font-bold resize-none" />
                         </div>
                         <button className="w-full bg-brand-gradient text-white py-5 rounded-2xl font-black text-xs uppercase tracking-widest shadow-xl shadow-brand-green/20 hover:opacity-90 active:scale-95 transition-all">
                            {lang === 'en' ? 'Send Inqury' : 'অনুরোধ পাঠান'}
                         </button>
                      </form>
                   </div>
                </div>
              </section>
            )}
                        
                        <div className="flex flex-wrap gap-4 mt-8">
                          <label className="inline-flex items-center gap-3 px-10 py-5 bg-brand-green text-white rounded-2xl font-black text-sm uppercase tracking-widest shadow-xl shadow-brand-green/20 hover:scale-105 active:scale-95 transition-all cursor-pointer">
                            <Plus size={20} />
                            Select Excel File
                            <input type="file" accept=".xlsx, .xls" onChange={handleBulkImport} className="hidden" />
                          </label>
                          <button 
                            onClick={downloadSampleExcel}
                            className="inline-flex items-center gap-3 px-10 py-5 bg-white text-brand-green border-2 border-brand-green/20 rounded-2xl font-black text-sm uppercase tracking-widest hover:bg-brand-green/5 transition-all"
                          >
                            <Download size={20} />
                            {t.downloadSample}
                          </button>
                        </div>
                      </div>

                      <div className="p-10 bg-gray-50 rounded-[2.5rem] border border-gray-100 group">
                         <div className="w-16 h-16 bg-brand-yellow/10 rounded-2xl flex items-center justify-center mb-6 text-brand-yellow">
                           <Building2 size={32} />
                         </div>
                         <h3 className="text-2xl font-black text-gray-900 mb-4">System Stats</h3>
                         <div className="grid grid-cols-2 gap-4">
                            <div className="bg-white p-6 rounded-2xl border border-gray-100">
                               <div className="text-3xl font-black text-brand-green">{students.length}</div>
                               <div className="text-[10px] font-black text-gray-400 uppercase tracking-widest mt-1">Total Students</div>
                            </div>
                            <div className="bg-white p-6 rounded-2xl border border-gray-100">
                               <div className="text-3xl font-black text-brand-yellow">{registeredStudents.length}</div>
                               <div className="text-[10px] font-black text-gray-400 uppercase tracking-widest mt-1">Portal Users</div>
                            </div>
                         </div>
                      </div>
                    </div>

                    <div className="p-10 bg-gray-50 rounded-[2.5rem] border border-gray-100">
                      <h3 className="text-2xl font-black text-gray-900 mb-8">{t.semesterDistribution}</h3>
                      <div className="h-80 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={semesterStats}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#eee" />
                            <XAxis 
                              dataKey="name" 
                              axisLine={false} 
                              tickLine={false} 
                              tick={{ fill: '#94a3b8', fontSize: 12, fontWeight: 600 }}
                            />
                            <YAxis 
                              axisLine={false} 
                              tickLine={false} 
                              tick={{ fill: '#94a3b8', fontSize: 12, fontWeight: 600 }}
                            />
                            <Tooltip 
                              cursor={{ fill: 'rgba(17, 101, 48, 0.05)' }}
                              contentStyle={{ 
                                borderRadius: '1rem', 
                                border: 'none', 
                                boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.1)',
                                padding: '1rem'
                              }}
                            />
                            <Bar 
                              dataKey="count" 
                              fill="#116530" 
                              radius={[8, 8, 0, 0]}
                              barSize={40}
                            >
                              {semesterStats.map((entry, index) => (
                                <Cell key={`cell-${index}`} fill={entry.count > 0 ? '#116530' : '#e2e8f0'} />
                              ))}
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  </div>
                </div>
              </section>
            )}
          </motion.div>
        </AnimatePresence>
      </main>

      {/* Footer */}
      <footer className="relative bg-[#1a1c1e] text-[#ccc] pt-16 pb-10 overflow-hidden">
        <div className="absolute top-0 right-0 w-1/2 h-full bg-white/[0.02] skew-x-12 translate-x-1/2"></div>
        <div className="max-w-7xl mx-auto grid md:grid-cols-3 gap-12 relative z-10 px-4">
          {/* Column 1: Brand */}
          <div className="space-y-6">
            <div className="flex items-center gap-3">
              <div className="bg-brand-green p-2.5 rounded-xl text-white shadow-2xl">
                <Leaf size={24} />
              </div>
              <h3 className="text-xl font-black tracking-tighter text-white uppercase italic font-display">
                {lang === 'en' ? 'Food Tech CNPI' : 'ফুড টেকনোলজি সিএনপিআই'}
              </h3>
            </div>
            <p className="font-medium leading-relaxed italic text-sm border-l-4 border-brand-green pl-6 text-[#ccc]/80">
              {t.footerMsg}
            </p>
            <div className="flex items-center gap-4 pt-2">
              {user ? (
                <div className="flex items-center gap-4 bg-white/5 p-3 rounded-2xl border border-white/10">
                  <img src={user.photoURL || `https://ui-avatars.com/api/?name=${encodeURIComponent(user.displayName || 'User')}&background=116530&color=fff`} className="w-8 h-8 rounded-full border-2 border-brand-yellow" />
                  <div className="flex flex-col">
                    <span className="text-[10px] text-white font-black uppercase tracking-widest">{user.displayName}</span>
                    <button onClick={logout} className="text-[10px] font-bold text-brand-yellow hover:underline flex items-center gap-1.5 mt-1">
                      <LogOut size={12} /> Sign Out
                    </button>
                  </div>
                </div>
              ) : (
                <button 
                  onClick={login} 
                  className="px-6 py-3 bg-white/5 hover:bg-white/10 rounded-xl text-[10px] font-black text-white uppercase tracking-widest flex items-center gap-2 transition-all border border-white/10"
                >
                  <LogIn size={14} /> Admin Access
                </button>
              )}
            </div>
          </div>
          
          {/* Column 2: Quick Links */}
          <div>
            <h3 className="text-xl font-black mb-10 text-white uppercase tracking-wider flex items-center gap-3">
              <div className="w-2 h-8 bg-brand-yellow rounded-full"></div>
              {t.quickLinks}
            </h3>
            <ul className="grid grid-cols-1 gap-4">
              <QuickFooterLink text={t.bteb} href="https://bteb.gov.bd/" external />
              <QuickFooterLink text={t.dte} href="https://techedu.gov.bd/" external />
              <QuickFooterLink text={t.tmed} href="https://tmed.gov.bd/" external />
              <QuickFooterLink text={t.cnpi} href="http://cnpi.chapainawabganj.gov.bd/" external />
              <QuickFooterLink text={t.studentPortal} onClick={() => setActiveTab('student-portal')} />
              <QuickFooterLink text={t.officeStaff} onClick={() => setActiveTab('staff')} />
            </ul>
          </div>

          {/* Column 3: Contact */}
          <div>
             <h3 className="text-xl font-black mb-10 text-white uppercase tracking-wider flex items-center gap-3">
               <div className="w-2 h-8 bg-brand-green rounded-full"></div>
               {t.contact}
             </h3>
             <div className="space-y-6">
               <ContactUnit icon={MapPin} text={t.location} />
               <ContactUnit icon={Phone} text="+880 1737256030" />
               <ContactUnit icon={Mail} text="food.chapaipoly@gmail.com" />
               
               <div className="flex gap-4 pt-4">
                 <SocialIcon icon={Facebook} />
                 <SocialIcon icon={Youtube} />
                 <SocialIcon icon={Linkedin} />
               </div>
             </div>
          </div>
        </div>
        
        <div className="max-w-7xl mx-auto mt-24 pt-12 border-t border-white/5 flex flex-col md:flex-row justify-between items-center gap-10 px-4">
          <p className="text-[11px] font-black text-[#888] uppercase tracking-[0.2em] text-center md:text-left transition-colors hover:text-[#bbb]">
            © {new Date().getFullYear()} {t.dept}, CNPI. {t.rights} <br className="md:hidden" />
            Designed By <a href="http://nasim.ami.bd" target="_blank" rel="noopener noreferrer" className="text-brand-green hover:text-brand-yellow transition-all font-black border-b border-brand-green/30 px-1 inline-block">{t.developer}</a>
          </p>
          <div className="flex flex-wrap justify-center gap-8 font-black text-[10px] text-[#666] uppercase tracking-[0.2em]">
            <a href="#" className="hover:text-white transition-colors">Privacy Policy</a>
            <a href="#" className="hover:text-white transition-colors">Terms of Service</a>
            <a href="#" className="hover:text-white transition-colors">Cookie Policy</a>
          </div>
        </div>
      </footer>

      {/* Back to Top */}
      <AnimatePresence>
        {showBackToTop && (
          <motion.button
            initial={{ opacity: 0, scale: 0.8 }}
            animate={{ opacity: 1, scale: 1 }}
            exit={{ opacity: 0, scale: 0.8 }}
            onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}
            className="fixed bottom-10 right-10 z-50 bg-brand-green text-white p-5 rounded-3xl shadow-2xl hover:bg-brand-green/90 transition-all hover:-translate-y-2 active:scale-95 group border-2 border-brand-yellow/30"
          >
            <ChevronUp size={28} className="group-hover:animate-bounce" />
          </motion.button>
        )}
      </AnimatePresence>
    </div>
  );
}

function EnhancedFeatureItem({ icon: Icon, title, desc, color }: { icon: any; title: string; desc: string; color: 'green' | 'yellow' | 'slate' }) {
  const colorClasses = {
    green: 'bg-emerald-50 text-emerald-600 border-emerald-100 group-hover:bg-emerald-600',
    yellow: 'bg-amber-50 text-amber-600 border-amber-100 group-hover:bg-amber-600',
    slate: 'bg-slate-50 text-slate-600 border-slate-100 group-hover:bg-slate-600'
  };

  return (
    <div className="bg-white border border-slate-100 p-8 rounded-3xl flex flex-col gap-6 transition-all hover:shadow-2xl hover:-translate-y-2 group">
      <div className={`w-14 h-14 rounded-2xl flex items-center justify-center transition-all duration-300 border ${colorClasses[color]} group-hover:text-white group-hover:shadow-lg shadow-sm`}>
        <Icon size={28} />
      </div>
      <div>
        <h4 className="font-black text-slate-900 text-lg mb-2 tracking-tight">{title}</h4>
        <p className="text-sm text-slate-500 font-medium leading-relaxed">{desc}</p>
      </div>
    </div>
  );
}

function FeatureItem({ icon: Icon, label }: { icon: any; label: string }) {
  return (
    <motion.div 
      whileHover={{ y: -5 }}
      className="p-6 bento-card flex flex-col items-center text-center gap-3 group"
    >
      <div className="p-3 bg-brand-accent text-brand-green rounded-xl group-hover:bg-brand-gradient group-hover:text-white transition-all duration-500 shadow-sm">
        <Icon size={20} />
      </div>
      <p className="text-[11px] font-black text-slate-800 uppercase tracking-widest leading-tight font-display">{label}</p>
    </motion.div>
  );
}

function SocialIcon({ icon: Icon }: { icon: any }) {
  return (
    <a href="#" className="w-12 h-12 rounded-xl border border-brand-green/10 bg-brand-green/5 flex items-center justify-center hover:bg-brand-green text-brand-green hover:text-white transition-all hover:scale-110 active:scale-95 group">
      <Icon size={20} />
    </a>
  );
}

function EnhancedSelect({ value, onChange, options, placeholder }: { value: string; onChange: (v: string) => void; options: string[]; placeholder: string }) {
  return (
    <div className="relative group">
       <select 
         value={value} 
         onChange={(e) => onChange(e.target.value)}
         className="w-full pl-6 pr-12 py-5 bg-gray-50 rounded-[1.5rem] appearance-none focus:ring-2 focus:ring-brand-green/20 font-black text-sm text-gray-700 border-none cursor-pointer tracking-tight outline-none"
       >
         <option value="">{placeholder}</option>
         {options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
       </select>
       <div className="absolute right-6 top-1/2 -translate-y-1/2 text-brand-green pointer-events-none group-focus-within:rotate-180 transition-transform">
         <ChevronDown size={18} />
       </div>
    </div>
  );
}

function ContactUnit({ icon: Icon, text }: { icon: any; text: string }) {
  return (
    <div className="flex gap-4 group items-center">
       <div className="p-3 bg-brand-green/5 text-brand-green rounded-xl transition-all group-hover:bg-brand-green group-hover:text-white shadow-sm shrink-0">
          <Icon size={18} />
       </div>
       <span className="text-xs font-bold text-gray-500 group-hover:text-gray-900 transition-colors leading-tight tracking-tight">{text}</span>
    </div>
  );
}

function QuickFooterLink({ text, href = "#", onClick, external }: { text: string; href?: string; onClick?: () => void; external?: boolean }) {
  return (
    <li>
      <a 
        href={href} 
        target={external ? "_blank" : undefined}
        rel={external ? "noopener noreferrer" : undefined}
        onClick={(e) => {
          if (onClick) {
            e.preventDefault();
            onClick();
            window.scrollTo({ top: 0, behavior: 'smooth' });
          }
        }}
        className="text-[#888] hover:text-brand-green transition-all flex items-center gap-4 group text-[11px] font-black uppercase tracking-widest leading-none"
      >
        <ChevronRight size={14} className="text-brand-green group-hover:translate-x-1 transition-transform" />
        {text}
      </a>
    </li>
  );
}

function LabPoint({ text }: { text: string }) {
  return (
    <li className="flex items-center gap-4 text-sm font-black text-gray-600 bg-white p-5 rounded-[1.5rem] border border-brand-green/10 shadow-sm hover:translate-x-2 transition-all">
      <div className="w-3 h-3 bg-brand-yellow rounded-full shadow-lg shadow-brand-yellow/40"></div>
      {text}
    </li>
  );
}

function ProfileEditor({ profile, onSave, lang, t }: { profile: StudentProfile, onSave: (p: StudentProfile) => void, lang: string, t: any }) {
  const [activeSubTab, setActiveSubTab] = useState('basic');
  const [localProfile, setLocalProfile] = useState<StudentProfile>(profile);
  const [isUploading, setIsUploading] = useState(false);
  const [errors, setErrors] = useState<Record<string, string>>({});

  const validateField = (field: string, value: string) => {
    let error = '';
    if (field === 'email') {
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (value && !emailRegex.test(value)) {
        error = 'Invalid email address';
      }
    } else if (field === 'mobile') {
      const mobileRegex = /^01[3-9]\d{8}$/;
      if (value && !mobileRegex.test(value)) {
        error = 'Invalid BD mobile number (11 digits, starts with 01)';
      }
    } else if (field === 'roll' || field === 'registration') {
      if (value && !/^\d+$/.test(value)) {
        error = 'Must be numbers only';
      }
    } else if (field === 'nid' || field === 'birthReg') {
      if (value && !/^\d{10,17}$/.test(value)) {
        error = 'Must be 10-17 digits';
      }
    } else if (field === 'dob') {
        if (value) {
            const date = new Date(value);
            if (isNaN(date.getTime())) {
                error = 'Invalid date';
            }
        }
    }
    return error;
  };

  const clearBasicInfo = () => {
    if (window.confirm("Are you sure you want to clear all Basic Information fields? This action cannot be undone.")) {
      const clearedBasic = {
        ...localProfile.basic,
        name: '',
        roll: '',
        registration: '',
        semester: '',
        shift: '',
        group: '',
        session: '',
        mobile: '',
        dob: '',
        birthReg: '',
        nid: '',
        image: ''
      };
      setLocalProfile({
        ...localProfile,
        basic: clearedBasic
      });
      // Clear errors for basic section
      const newErrors = { ...errors };
      Object.keys(newErrors).forEach(key => {
        if (key.startsWith('basic.')) delete newErrors[key];
      });
      setErrors(newErrors);
    }
  };

  const handleImageUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Check file size (limit to 2MB for profile pictures)
    if (file.size > 2 * 1024 * 1024) {
      alert("Image size should be less than 2MB.");
      return;
    }

    setIsUploading(true);
    try {
      const storageRef = ref(storage, `profiles/${profile.userId}_${Date.now()}_${file.name}`);
      console.log("Starting upload to:", storageRef.fullPath);
      
      const metadata = {
        contentType: file.type,
      };

      await uploadBytes(storageRef, file, metadata);
      const url = await getDownloadURL(storageRef);
      handleChange('basic', 'image', url);
    } catch (error: any) {
      console.error("Detailed upload error:", error);
      let errorMessage = "Failed to upload image.";
      
      if (error.code === 'storage/retry-limit-exceeded') {
        errorMessage = "Connection timeout. Please check your internet or try a smaller file. Ensure Firebase Storage is enabled in your console.";
      } else if (error.code === 'storage/unauthorized') {
        errorMessage = "Permission denied. Please make sure you are logged in correctly.";
      } else if (error.message) {
        errorMessage += " " + error.message;
      }
      
      alert(errorMessage);
    } finally {
      setIsUploading(false);
    }
  };

  const handleChange = (section: string, field: string, value: any, subSection?: string) => {
    const errorKey = subSection ? `${subSection}.${field}` : `${section}.${field}`;
    const error = validateField(field, value);
    
    setErrors(prev => ({
      ...prev,
      [errorKey]: error
    }));

    const updated = { ...localProfile } as any;
    if (subSection) {
      if (!updated[section][subSection]) updated[section][subSection] = {};
      updated[section][subSection][field] = value;
    } else {
      if (!updated[section]) updated[section] = {};
      updated[section][field] = value;
    }
    setLocalProfile({ ...updated });
  };

  const handleSave = () => {
    // Final validation check
    const hasErrors = Object.values(errors).some(err => err !== '');
    if (hasErrors) {
      alert("Please fix validation errors before saving.");
      return;
    }
    onSave(localProfile);
  };

  const addEducation = () => {
    setLocalProfile({
      ...localProfile,
      education: [...localProfile.education, { institution: '', exam: '', subject: '', board: '', result: '', year: '' }]
    });
  };

  const removeEducation = (index: number) => {
    const updated = [...localProfile.education];
    updated.splice(index, 1);
    setLocalProfile({ ...localProfile, education: updated });
  };

  const updateEducation = (index: number, field: string, value: string) => {
    const updated = [...localProfile.education];
    const item = { ...updated[index], [field]: value };
    updated[index] = item;
    setLocalProfile({ ...localProfile, education: updated });
  };

  const tabs = [
    { id: 'basic', label: t.basicInfo },
    { id: 'education', label: t.eduInfo },
    { id: 'parents', label: t.parentInfo },
    { id: 'address', label: t.addressInfo },
    { id: 'pro', label: t.proInfo },
    { id: 'fin', label: t.finInfo }
  ];

  return (
    <div className="bg-white rounded-[3rem] shadow-2xl border border-gray-100 overflow-hidden">
      <div className="flex overflow-x-auto border-b border-gray-100 scrollbar-hide">
        {tabs.map(tab => (
          <button
            key={tab.id}
            onClick={() => setActiveSubTab(tab.id)}
            className={`px-8 py-6 text-xs font-black uppercase tracking-widest whitespace-nowrap transition-all ${activeSubTab === tab.id ? 'text-brand-green bg-brand-green/5 border-b-4 border-brand-green' : 'text-gray-400 hover:bg-gray-50'}`}
          >
            {tab.label}
          </button>
        ))}
      </div>

      <div className="p-10 space-y-10">
        {activeSubTab === 'basic' && (
          <div className="space-y-8">
            <div className="flex justify-between items-center mb-4">
              <h3 className="text-xl font-black text-gray-900 tracking-tight">Basic Information</h3>
              <button 
                onClick={clearBasicInfo}
                className="text-[10px] font-black uppercase tracking-widest text-red-500 hover:text-red-700 transition-colors flex items-center gap-2"
              >
                <Trash2 size={14} />
                Clear All Fields
              </button>
            </div>
            <div className="grid md:grid-cols-2 gap-8">
            {[
              { label: 'Full Name', field: 'name' },
              { label: 'Roll No', field: 'roll' },
              { label: 'Registration No', field: 'registration' },
              { label: 'Mobile No', field: 'mobile' },
              { label: 'MIS ID', field: 'misId' },
              { label: 'Semester', field: 'semester' },
              { label: 'Shift', field: 'shift' },
              { label: 'Group', field: 'group' },
              { label: 'Session', field: 'session' },
              { label: 'Date of Birth', field: 'dob', type: 'date' },
              { label: 'Birth Registration No', field: 'birthReg' },
              { label: 'NID No', field: 'nid' },
            ].map(item => (
              <div key={item.field} className="space-y-2">
                <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">{item.label}</label>
                <input 
                  type={item.type || 'text'}
                  value={(localProfile.basic as any)[item.field] || ''}
                  onChange={(e) => handleChange('basic', item.field, e.target.value)}
                  className={`w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 outline-none font-bold ${errors[`basic.${item.field}`] ? 'ring-2 ring-red-500' : 'focus:ring-brand-green/20'}`} 
                />
                {errors[`basic.${item.field}`] && (
                  <p className="text-[10px] font-bold text-red-500 ml-4">{errors[`basic.${item.field}`]}</p>
                )}
              </div>
            ))}
          </div>
            
          <div className="space-y-2">
              <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Profile Image</label>
              <div className="flex flex-col gap-4">
                <div className="flex items-center gap-6">
                  {localProfile.basic.image && (
                    <img src={localProfile.basic.image} className="w-20 h-20 rounded-2xl object-cover border-2 border-brand-green/10" alt="Preview" />
                  )}
                  <label className={`flex-1 flex items-center justify-center gap-3 px-8 py-4 bg-white border-2 border-dashed border-gray-200 rounded-2xl cursor-pointer hover:border-brand-green/30 transition-all ${isUploading ? 'opacity-50 pointer-events-none' : ''}`}>
                    <Upload size={20} className="text-gray-400" />
                    <span className="text-sm font-bold text-gray-500">{isUploading ? 'Uploading...' : 'Upload Image'}</span>
                    <input type="file" accept="image/*" onChange={handleImageUpload} className="hidden" />
                  </label>
                </div>
                <div className="space-y-2">
                  <label className="text-[8px] font-black uppercase tracking-widest text-gray-300 ml-4">Or Image URL</label>
                  <input 
                    type="text"
                    value={localProfile.basic.image || ''}
                    onChange={(e) => handleChange('basic', 'image', e.target.value)}
                    placeholder="https://..."
                    className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold text-xs" 
                  />
                </div>
              </div>
            </div>
          </div>
        )}

        {activeSubTab === 'education' && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
               <h3 className="text-xl font-black text-gray-900 tracking-tight">Academic History</h3>
               <button onClick={addEducation} className="bg-brand-green text-white p-3 rounded-xl hover:scale-110 active:scale-95 transition-all shadow-lg shadow-brand-green/10">
                 <Plus size={20} />
               </button>
            </div>
            {localProfile.education.map((edu, i) => (
              <div key={i} className="p-8 bg-gray-50 rounded-[2rem] border border-gray-100 relative group">
                <button onClick={() => removeEducation(i)} className="absolute -top-3 -right-3 p-2 bg-red-500 text-white rounded-lg opacity-0 group-hover:opacity-100 transition-all shadow-lg">
                   <X size={14} />
                </button>
                <div className="grid md:grid-cols-2 lg:grid-cols-3 gap-6">
                   {[
                     { label: 'Institution', field: 'institution' },
                     { label: 'Exam Name', field: 'exam' },
                     { label: 'Group / Subject', field: 'subject' },
                     { label: 'Board', field: 'board' },
                     { label: 'GPA / CGPA', field: 'result' },
                     { label: 'Passing Year', field: 'year' },
                   ].map(item => (
                      <div key={item.field} className="space-y-2">
                        <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-2">{item.label}</label>
                        <input 
                          value={(edu as any)[item.field]}
                          onChange={(e) => updateEducation(i, item.field, e.target.value)}
                          className="w-full px-6 py-3 bg-white rounded-xl border border-gray-200 focus:ring-2 focus:ring-brand-green/20 outline-none font-bold text-sm" 
                        />
                      </div>
                   ))}
                </div>
              </div>
            ))}
          </div>
        )}

        {activeSubTab === 'parents' && (
          <div className="space-y-16">
            <div className="grid md:grid-cols-2 gap-16">
              {['father', 'mother'].map(parent => (
                <div key={parent} className="space-y-8">
                  <h3 className="text-xl font-black text-gray-900 tracking-tight capitalize">{parent === 'father' ? "Father's" : "Mother's"} Information</h3>
                  <div className="space-y-6">
                     {[
                       { label: 'Name', field: 'name' },
                       { label: 'Occupation', field: 'occupation' },
                       { label: 'Mobile No', field: 'mobile' },
                       { label: 'NID No', field: 'nid' },
                       { label: 'Date of Birth', field: 'dob', type: 'date' },
                     ].map(item => (
                        <div key={item.field} className="space-y-2">
                          <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">{item.label}</label>
                          <input 
                            type={item.type || 'text'}
                            value={(localProfile.parents as any)[parent][item.field] || ''}
                            onChange={(e) => handleChange('parents', item.field, e.target.value, parent)}
                            className={`w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 outline-none font-bold ${errors[`${parent}.${item.field}`] ? 'ring-2 ring-red-500' : 'focus:ring-brand-green/20'}`} 
                          />
                          {errors[`${parent}.${item.field}`] && (
                            <p className="text-[10px] font-bold text-red-500 ml-4">{errors[`${parent}.${item.field}`]}</p>
                          )}
                        </div>
                     ))}
                  </div>
                </div>
              ))}
            </div>

            <div className="pt-10 border-t border-gray-100">
              <h3 className="text-xl font-black text-gray-900 tracking-tight mb-8">Guardian Information</h3>
              <div className="grid md:grid-cols-2 gap-8">
                 {[
                   { label: 'Guardian Name', field: 'name' },
                   { label: 'Relation', field: 'relation' },
                   { label: 'Mobile No', field: 'mobile' },
                   { label: 'Occupation', field: 'occupation' },
                 ].map(item => (
                    <div key={item.field} className="space-y-2">
                      <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">{item.label}</label>
                      <input 
                        value={(localProfile.parents as any).guardian?.[item.field] || ''}
                        onChange={(e) => handleChange('parents', item.field, e.target.value, 'guardian')}
                        className={`w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 outline-none font-bold ${errors[`guardian.${item.field}`] ? 'ring-2 ring-red-500' : 'focus:ring-brand-green/20'}`} 
                      />
                      {errors[`guardian.${item.field}`] && (
                        <p className="text-[10px] font-bold text-red-500 ml-4">{errors[`guardian.${item.field}`]}</p>
                      )}
                    </div>
                 ))}
              </div>
            </div>
          </div>
        )}

        {activeSubTab === 'address' && (
          <div className="grid md:grid-cols-2 gap-8">
            <div className="space-y-2">
               <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Current Address</label>
               <textarea 
                  rows={4}
                  value={localProfile.address.current}
                  onChange={(e) => handleChange('address', 'current', e.target.value)}
                  className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold resize-none" 
               />
            </div>
            <div className="space-y-2">
               <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">Permanent Address</label>
               <textarea 
                  rows={4}
                  value={localProfile.address.permanent}
                  onChange={(e) => handleChange('address', 'permanent', e.target.value)}
                  className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold resize-none" 
               />
            </div>
          </div>
        )}

        {activeSubTab === 'pro' && (
          <div className="space-y-8">
            {[
              { label: 'Career Objective', field: 'objective' },
              { label: 'Professional Experience', field: 'experience' },
              { label: 'Extra Curricular & Certificates', field: 'certificates' },
              { label: 'Language Skills', field: 'languages' },
            ].map(item => (
              <div key={item.field} className="space-y-2">
                <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">{item.label}</label>
                <textarea 
                  rows={3}
                  value={(localProfile.professional as any)[item.field]}
                  onChange={(e) => handleChange('professional', item.field, e.target.value)}
                  className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold resize-none" 
                />
              </div>
            ))}
          </div>
        )}

        {activeSubTab === 'fin' && (
          <div className="grid md:grid-cols-2 gap-8">
            {[
              { label: 'Bank Name', field: 'bankName' },
              { label: 'Account Holder Name', field: 'accName' },
              { label: 'Account Number', field: 'accNo' },
              { label: 'Associated NID', field: 'accNid' },
              { label: 'bKash Number', field: 'bkash' },
              { label: 'Rocket Number', field: 'rocket' },
            ].map(item => (
              <div key={item.field} className="space-y-2">
                <label className="text-[10px] font-black uppercase tracking-widest text-gray-400 ml-4">{item.label}</label>
                <input 
                  value={(localProfile.financial as any)[item.field]}
                  onChange={(e) => handleChange('financial', item.field, e.target.value)}
                  className="w-full px-8 py-4 bg-gray-50 rounded-2xl focus:ring-2 focus:ring-brand-green/20 outline-none font-bold" 
                />
              </div>
            ))}
          </div>
        )}

        <div className="pt-10 border-t border-gray-100 flex justify-end">
           <button 
             onClick={handleSave}
             disabled={isUploading}
             className={`px-12 py-5 bg-brand-green text-white rounded-2xl font-black text-sm uppercase tracking-widest shadow-xl shadow-brand-green/20 hover:scale-105 active:scale-95 transition-all ${isUploading ? 'opacity-50 cursor-not-allowed' : ''}`}
           >
             {isUploading ? 'Uploading...' : 'Save Changes'}
           </button>
        </div>
      </div>
    </div>
  );
}

function printCV(profile: StudentProfile | null) {
  if (!profile) return;
  const printWindow = window.open('', '', 'width=800,height=1000');
  if (!printWindow) return;

  const eduRows = profile.education.map(edu => `
    <tr>
      <td>${edu.exam}</td>
      <td>${edu.institution}</td>
      <td>${edu.board}</td>
      <td>${edu.year}</td>
      <td>${edu.result}</td>
    </tr>
  `).join('');

  printWindow.document.write(`
    <html>
      <head>
        <title>CV - ${profile.basic.name}</title>
        <style>
          @page { size: A4; margin: 15mm; }
          body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.5; color: #333; margin: 0; padding: 0; }
          .container { width: 100%; }
          .header { display: flex; justify-content: space-between; align-items: start; border-bottom: 2px solid #116530; padding-bottom: 20px; margin-bottom: 30px; }
          .header-info h1 { margin: 0; color: #116530; font-size: 28px; text-transform: uppercase; }
          .header-info p { margin: 5px 0; color: #666; font-size: 14px; }
          .photo { width: 120px; height: 120px; border: 3px solid #f2f2f2; object-fit: cover; }
          
          h2 { font-size: 16px; text-transform: uppercase; border-bottom: 1px solid #eee; padding-bottom: 5px; margin-top: 25px; color: #116530; }
          
          .section { margin-bottom: 20px; }
          .grid { display: grid; grid-template-columns: 150px 1fr; gap: 10px; font-size: 13px; }
          .label { font-weight: bold; color: #555; }
          
          table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 12px; }
          th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
          th { background: #f9f9f9; }
          
          .objective { font-size: 13px; font-style: italic; color: #444; }
        </style>
      </head>
      <body onload="window.print()">
        <div class="container">
          <div class="header">
            <div class="header-info">
              <h1>${profile.basic.name}</h1>
              <p>Roll: ${profile.basic.roll} | Reg: ${profile.basic.registration} | Session: ${profile.basic.session || 'N/A'}</p>
              <p>Email: ${profile.basic.email}</p>
              <p>Mobile: ${profile.basic.mobile}</p>
              <p>Semester: ${profile.basic.semester} | Shift: ${profile.basic.shift} | Group: ${profile.basic.group}</p>
            </div>
            <img class="photo" src="${profile.basic.image || 'https://ui-avatars.com/api/?name='+encodeURIComponent(profile.basic.name)}">
          </div>

          <div class="section">
            <h2>Career Objective</h2>
            <p class="objective">${profile.professional.objective || 'Not specified'}</p>
          </div>

          <div class="section">
            <h2>Educational Qualifications</h2>
            <table>
              <thead>
                <tr>
                  <th>Examination</th>
                  <th>Institution</th>
                  <th>Board</th>
                  <th>Year</th>
                  <th>Result</th>
                </tr>
              </thead>
              <tbody>${eduRows || '<tr><td colspan="5" style="text-align:center">No data available</td></tr>'}</tbody>
            </table>
          </div>

          <div class="section">
            <h2>Personal Information</h2>
            <div class="grid">
              <span class="label">Father's Name:</span><span>${profile.parents.father.name}</span>
              <span class="label">Mother's Name:</span><span>${profile.parents.mother.name}</span>
              ${profile.parents.guardian?.name ? `
                <span class="label">Guardian Name:</span><span>${profile.parents.guardian.name} (${profile.parents.guardian.relation})</span>
                <span class="label">Guardian Mobile:</span><span>${profile.parents.guardian.mobile}</span>
              ` : ''}
              <span class="label">Date of Birth:</span><span>${profile.basic.dob}</span>
              <span class="label">NID No:</span><span>${profile.basic.nid || '-'}</span>
              <span class="label">Birth Reg No:</span><span>${profile.basic.birthReg || '-'}</span>
              <span class="label">MIS ID:</span><span>${profile.basic.misId || '-'}</span>
            </div>
          </div>

          <div class="section">
            <h2>Present Address</h2>
            <p style="font-size: 13px;">${profile.address.current || 'Not provided'}</p>
          </div>

          <div class="section">
            <h2>Permanent Address</h2>
            <p style="font-size: 13px;">${profile.address.permanent || 'Not provided'}</p>
          </div>

          <div class="section" style="margin-top: 50px;">
            <div style="float: right; text-align: center; width: 200px;">
              <div style="border-top: 1px solid #000; padding-top: 5px; font-size: 13px; font-weight: bold;">
                Applicant Signature
              </div>
            </div>
          </div>
        </div>
      </body>
    </html>
  `);
  printWindow.document.close();
}
