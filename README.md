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
> ▶️ Click the thumbnail above to watch the full SAMS project demo on YouTube. <br>
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

---

## 🖼️ Screenshots

Below are complete interface screenshots from the SAMS application, covering all key modules and workflows — including login, admin controls, student management, attendance marking, user dashboard, and detailed reports.

<table>
  <tr>
    <td><img src="https://i.ibb.co/6d5QYJn/z1.png" width="400"/></td>
    <td><img src="https://i.ibb.co/DHLSHb1K/z2.png" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://i.ibb.co/ZpBcqQZv/z3.png" width="400"/></td>
    <td><img src="https://i.ibb.co/SwBCHzMt/z4.png" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://i.ibb.co/KZvS8Y6/z5.png" width="400"/></td>
    <td><img src="https://i.ibb.co/0RpxMCmK/z6.png" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://i.ibb.co/nN64CcMr/z7.png" width="400"/></td>
    <td><img src="https://i.ibb.co/8ZBW90X/z8.png" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://i.ibb.co/SXyMMLLz/z9.png" width="400"/></td>
    <td><img src="https://i.ibb.co/Jj64nZg8/z10.png" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://i.ibb.co/cRvSGWg/z11.png" width="400"/></td>
    <td><img src="https://i.ibb.co/Gf0779Rr/z12.png" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://i.ibb.co/NgZ89Mr4/z13.png" width="400"/></td>
    <td><img src="https://i.ibb.co/GvNC23zk/z14.png" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://i.ibb.co/nq8rPsb2/z15.png" width="400"/></td>
    <td><img src="https://i.ibb.co/XxxgBpgJ/z16.png" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://i.ibb.co/pvMmx527/z17.png" width="400"/></td>
    <td><img src="https://i.ibb.co/v4Yhn6Nd/z18.png" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://i.ibb.co/SwPJY1KP/z19.png" width="400"/></td>
    <td><img src="https://i.ibb.co/LDKNC4Ph/z20.png" width="400"/></td>
  </tr>
</table>

> 📌 This gallery includes all major forms and features of the project to give a full visual understanding of how SAMS works.

---

## 🛡️ Installation

Follow these steps to install and run the SAMS application on your Windows system:

1. **Download the Installer**
   - Go to the [Releases](https://github.com/prashanth-kumar-g/student-attendance-management-system/releases/tag/v1.0.0/SAMS_Setup.exe) section.
   - Download the latest version of `SAMS_Setup.exe`.

2. **Run the Installer**
   - Double-click `SAMS_Setup.exe`.
   - It will install the application, SQL Server Express (if not already installed), and restore the SAMS database automatically.

3. **Launch the Application**
   - Use the desktop shortcut or Start menu to open the app.
   - Log in using the default admin credentials and begin exploring the application's features.

4. **Need Help?**
   - 📺 Refer to the [Demo Video](#demo-video) above for a full walkthrough of installation and usage.

> 💡 No manual database setup is required — everything is bundled with the installer for a seamless experience.

---

## 🧑‍💻 Developer Notes

If you are a developer and wish to rebuild or customize the SAMS installer, complete instructions are available in:

📁 [`Package/RebuildInstructions.md`](Package/RebuildInstructions.md)

This includes:
- How to clone the repo
- Where to download the full external installer package
- Folder structure required to compile
- Steps to build the `SAMS_Setup.exe` using Inno Setup

> ⚠️ End users do not need this. This is only for those who want to audit or regenerate the installer manually.

---

## 🤝 Contributing

Contributions are welcome!

If you'd like to improve this project, fix issues, or suggest new features:

1. Fork the repository
2. Create a new branch for your changes
3. Submit a pull request with a clear explanation

You can also open issues for bugs, ideas, or questions.  
Thank you for helping make SAMS better!

---

## 📜 License

This project is licensed under the [MIT License](LICENSE).

You are free to use and modify this project for personal or educational purposes.  
Please see the `LICENSE` file for full license terms.

---

## ©️ Copyrights

© 2025 Prashant Kumar G. All rights reserved.

This project, including its source code, design, and associated materials, is the original work of the author.  
Unauthorized copying, modification, or commercial use is not permitted without prior written consent.

---
