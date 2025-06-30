## 📦 SAMS Installer Package (Development Files)

This file contains instructions for developers who wish to rebuild the SAMS Setup installer (`SAMS_Setup.exe`) using the original Inno Setup script (`sams_installer.iss`).

To keep this GitHub repository clean and fast to clone, the full installer package (345 MB) is hosted externally.

🔗 Download the full package here:
[https://drive.google.com/drive/folders/1twyXTD3se0bn2g3-JdRF7pZ8PJwVoDI9?usp=sharing](https://drive.google.com/drive/folders/1twyXTD3se0bn2g3-JdRF7pZ8PJwVoDI9?usp=sharing)

🛠️ How to Rebuild the Installer:

1. **Clone this GitHub repository** to your local machine:

   ```bash
   git clone https://github.com/prashanth-kumar-g/student-attendance-management-system.git
   ```
2. Download **all files and folders** from the Google Drive link above.
3. Place them **inside the `Package/` folder** of this repository.
   Your structure should look like this:

   ```
   Package/
   ├── Bin/
   ├── Database/
   ├── Documents/
   ├── Images/
   ├── SQLExpressOffline/
   └── Student Attendance Management System.exe
   ```
4. Open `sams_installer.iss` in Inno Setup Compiler.
5. Verify that all paths in the script match the local structure.
6. Compile the script to generate your own `SAMS_Setup.exe`.

📌 Notes:

* End users do **not** need this package — they can directly download the final SAMS installer from the [GitHub Releases](https://github.com/prashanth-kumar-g/student-attendance-management-system/releases/tag/v1.0.0).
* This is intended for developers who wish to rebuild, audit, or customize the SAMS installer.
