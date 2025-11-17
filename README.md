# 🔧 Tooling Control App

<div align="center">

![Version](https://img.shields.io/badge/version-0.1.1-blue.svg)
![Electron](https://img.shields.io/badge/Electron-latest-47848F.svg?logo=electron)
![License](https://img.shields.io/badge/license-MIT-green.svg)

**A professional desktop application for managing tooling lifecycle, suppliers, and production tracking.**

[Features](#-features) • [Installation](#-installation) • [Usage](#-usage) • [Tech Stack](#-tech-stack) • [Screenshots](#-screenshots)

</div>

---

## ✨ Features

### 🏭 **Supplier Management**
- Comprehensive supplier tracking with real-time statistics
- Smart search and filtering capabilities
- Filter by expiration status (expired/expiring within 2 years)
- Visual indicators for critical tooling status

### 🛠️ **Tooling Lifecycle Control**
- Full lifecycle tracking from creation to expiration
- Production volume monitoring with visual progress bars
- Automatic expiration calculations and alerts
- Status management (Concluded, Under Analysis, Obsolete, etc.)
- Replacement chain tracking and visualization

### 📊 **Analytics Dashboard**
- Real-time statistics and KPIs
- Expiration forecasting
- Production volume analysis
- Visual charts and graphs

### 📎 **Attachment Management**
- Drag-and-drop file upload
- Organized by supplier and tooling item
- Quick access to documentation and drawings

### ✅ **Task Management**
- Built-in todo lists for each tooling item
- Track action items and follow-ups
- Mark tasks as complete

### 🎨 **Modern UI/UX**
- Clean, professional dark theme
- Expandable card-based layout
- Custom title bar with window controls
- Smooth animations and transitions
- Responsive design

---

## 📋 Requirements

- **Node.js** 16.x or higher
- **npm** 8.x or higher
- **Windows 10/11** (primary platform)

---

## 🚀 Installation

### Clone the repository
```bash
git clone <repository-url>
cd "Ferramental App"
```

### Install dependencies
```bash
npm install
```

### Run in development mode
```bash
npm start
```

### Build for production
```bash
npm run dist
```

The executable will be generated in the `dist` folder.

---

## 📖 Usage

### Starting the Application
1. Launch the application
2. Select a supplier from the sidebar
3. View and manage tooling items

### Managing Tooling
- **Add New**: Click the "+" button in the bottom right
- **Edit**: Click on any card to expand and edit fields
- **Delete**: Use the trash icon on expanded cards
- **Search**: Use the search icon to find specific items

### Filtering
- Click the **"⋮"** button next to "Suppliers"
- Enable filters for expired/expiring tooling
- Use the **X** badge to quickly clear filters

### Attachments
- Drag and drop files onto the attachment area
- Click the paperclip icon to view all attachments
- Organized automatically by supplier

---

## 🛠️ Tech Stack

| Technology | Purpose |
|------------|---------|
| **Electron** | Desktop application framework |
| **SQLite3** | Local database for data persistence |
| **Node.js** | Backend runtime |
| **HTML/CSS/JS** | Frontend interface |
| **Phosphor Icons** | Modern icon library |

---

## 📁 Project Structure

```
Ferramental App/
├── main.js                    # Electron main process
├── preload.js                 # Preload scripts for IPC
├── renderer.js                # Frontend logic
├── index.html                 # Main UI structure
├── style.css                  # Styling
├── package.json               # Dependencies and scripts
├── ferramental_database.db    # SQLite database (generated)
└── attachments/               # User attachments (generated)
```

---

## 🎯 Key Features in Detail

### Smart Filtering System
Filter tooling by expiration status:
- ⚠️ **Expired**: Already exceeded lifecycle
- ⏰ **Expiring Soon**: Within 1 year
- 📅 **Expiring**: Within 2 years

### Replacement Chain Tracking
- Visualize replacement relationships between tooling
- Track obsolescence and successor items
- Timeline view for replacement history

### Auto-Save Technology
- Real-time data synchronization
- Automatic field validation
- No manual save required

### Production Progress
- Visual progress bars showing lifecycle usage
- Percentage calculations
- Remaining quantity tracking

---

## 🎨 Screenshots

### Main Dashboard
*Clean interface with supplier sidebar and tooling cards*

### Analytics View
*Comprehensive statistics and charts*

### Filter Options
*Smart filtering for expired/expiring tooling*

---

## 🔧 Configuration

### Database
The application uses SQLite for local data storage. The database file `ferramental_database.db` is automatically created on first run.

### Attachments
Files are stored in the `attachments/` folder, organized by supplier name.

---

## 🤝 Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📝 License

This project is licensed under the MIT License.

---

## 🐛 Known Issues

- N/A

---

## 🗺️ Roadmap

- [ ] Export to Excel/PDF
- [ ] Multi-language support
- [ ] Cloud sync capabilities
- [ ] Advanced reporting features
- [ ] Email notifications for expiring tooling

---

## 👨‍💻 Author

**Rafael**

---

## 🙏 Acknowledgments

- Phosphor Icons for the beautiful icon set
- Electron community for excellent documentation
- SQLite for reliable local storage

---

<div align="center">

**Made with ❤️ for better tooling management**

⭐ Star this repo if you find it useful!

</div>
