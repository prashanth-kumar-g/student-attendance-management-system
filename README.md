# student-attendance-management-system

---

📋 Table of Contents

- [Overview](#-overview)
- [Demo Video](#-demo-video)
- [Tech Stack](#-tech-stack)
- [Features](#-features)
- [Screenshots](#%EF%B8%8F-screenshots)
- [Installation](#-installation)
- [Developer Notes](#-developer-notes)
- [Contributing](#-contributing)
- [License](#-license)
- [Copyrights](#%EF%B8%8F-copyrights)

---

## 🎯 Overview

✨ Student Attendance Management System (SAMS) is a Windows desktop application built with Visual Basic 6 and Microsoft SQL Server Express to help educational institutions manage attendance efficiently. It delivers a user-friendly interface with clear navigation for admins and students, guiding them through tasks like adding records or marking attendance in just a few clicks. It offers a one-click installer that sets up the application and database automatically, clear role-based dashboards for administrators and students, fast attendance marking, and detailed reporting. With its modular design and simple workflows, SAMS makes tracking attendance straightforward and reliable. Experience a seamless and complete attendance management solution designed for real-world academic use.

---

## 🎥 Demo Video

<div>
  <a href="https://www.youtube.com/watch?v=RSNttaydg-g" target="_blank">
    <img src="https://github.com/user-attachments/assets/6ee5a104-2fbd-4658-88c5-134f92fdc362" alt="SAMS Demo" width="480">
  </a>
</div>

> ▶️ Click the thumbnail above to watch the full SAMS project demo on YouTube.

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
    <td><img src="https://github.com/user-attachments/assets/621bce0c-3100-41aa-ad45-1c759d430449" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/3c2d4e66-8246-4c75-b7c0-5e8e29f1ab9a" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/1b01f45c-512a-49b6-9040-1a434ae48aab" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/4500481b-b330-4b77-a81c-a36f877d9f7b" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/cb28f65e-4357-43ed-ad2a-065a559dc359" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/14cd71f3-d19d-4d5c-8da9-32d8c90964ff" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/bb78baf4-640b-460a-a25d-eba02bd04f0f" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/d09f524d-b073-4529-8b07-8ef6590e1ff0" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/dc73acd4-bcae-409b-bd5b-e5bc1f0465de" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/776b1061-bb5a-4df1-9637-c23ce5e0b7b4" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/f45ae72c-351d-418b-ba19-74f3138709f5" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/909c632c-2f50-4d4e-8f4c-faebd41ce3b4" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/e8436dcd-ae19-429b-8167-7f4a8eea0893" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/69789a13-0816-4139-bcd3-b2bfc4f7ce29" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/521109b0-a857-45f2-b84f-39ab2cf1965e" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/e6bc2475-e3c0-4859-a352-25fe4dc4c8bb" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/fbafc665-1772-4d54-8268-5f6c94c3c2ce" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/bb208831-d428-44fa-8b92-ba332dbd8693" width="400"/></td>
  </tr>
  <tr>
    <td><img src="https://github.com/user-attachments/assets/7bf95da1-d54a-47ef-a067-645245fe9724" width="400"/></td>
    <td><img src="https://github.com/user-attachments/assets/6b66ee06-01c6-4a27-92dd-b195de0f6d15" width="400"/></td>
  </tr>
</table>

> 📌 This gallery includes all major forms and features of the project to give a full visual understanding of how SAMS works.

---

## 🚀 Installation

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
You may use and modify this code for personal or educational purposes—see `LICENSE` for full details.

---

## ©️ Copyrights

© 2025 Prashanth Kumar G. All rights reserved.  
Unauthorized commercial use or redistribution is prohibited without prior written consent.

---
