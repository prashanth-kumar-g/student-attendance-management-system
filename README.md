# student-attendance-management-system

---

📋 Table of Contents

- [Overview](#-overview)
- [Demo Video](#-demo-video)
- [Tech Stack](#-tech-stack)
- [Features](#-features)
- [Screenshots](#-screenshots)
- [Installation](#-installation)
- [Developer Notes](#-developer-notes)
- [Contributing](#-contributing)
- [License](#-license)
- [Copyrights](#-copyrights)

---

## 🧐 Overview

✨ Student Attendance Management System (SAMS) is a Windows desktop application built with Visual Basic 6 and Microsoft SQL Server Express to help educational institutions manage attendance efficiently. It delivers a user-friendly interface with clear navigation for admins and students, guiding them through tasks like adding records or marking attendance in just a few clicks. It offers a one-click installer that sets up the application and database automatically, clear role-based dashboards for administrators and students, fast attendance marking, and detailed reporting. With its modular design and simple workflows, SAMS makes tracking attendance straightforward and reliable. Experience a seamless and complete attendance management solution designed for real-world academic use.

---

## 🎥 Demo Video

<div>
  <a href="https://www.youtube.com/watch?v=RSNttaydg-g" target="_blank">
    <img src="https://i.ibb.co/rK4NBrZj/z5.png" alt="SAMS Demo" width="480">
  </a>
</div>
▶️ Click the thumbnail above to watch the full SAMS project demo on YouTube. <br>
<br>
📘 This demo video provides a complete walkthrough of the SAMS project, including installation, login flow, dashboard features, and attendance tracking. If you're new to the project or need a quick understanding of how it works, we recommend watching the video before proceeding with installation or development.

---

## 🛠 Tech Stack

The SAMS application is built using classic technologies suited for reliable offline desktop systems:

- **Programming Language:** Visual Basic 6.0  
- **Database:** Microsoft SQL Server Express 2022  
- **Installer:** Inno Setup (via `sams_installer.iss`)  
- **Architecture:** MDI (Multiple Document Interface) design  
- **Database Connectivity:** ADODB  
- **Setup Automation:** Batch script–based SQL database restoration during first run

---

## ✨ Features

SAMS is equipped with powerful and practical modules tailored for managing student attendance in educational institutions. Below is a breakdown of its key capabilities:

### 👨‍🏫 Admin Panel Features
- 🔐 Secure login authentication with validation for each admin
- 👨‍🎓 Add, edit, and manage student records with photo upload
- 👥 Add/edit/delete admin users with full details
- 🗓️ Mark daily attendance for all students from one interface
- 🔄 Update or delete existing records as needed

### 👨‍🎓 User Dashboard Features
- 📆 View detailed attendance reports by selected date
- 📊 Access per-student and full class attendance summaries
- 📄 Generate reports for: Student Details, Admin Details, Date-wise Attendance, Overall Attendance
- 🎨 Color-coded status labels for better readability (e.g., green for present, red for absent)
- 📋 Interactive report viewer with export-ready formatting

### ⚙️ Application Features
- 💾 First-launch automated database creation and restoration via batch script
- 📦 One-click installer that bundles SQL Server and runtime dependencies
- 🪟 MDI (Multi-Document Interface) for multitasking within the app
- 📚 Built-in help/documentation section (PDF)
- 🔒 Form tracking system to prevent multiple instances of the same window

