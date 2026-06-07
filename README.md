# 📄 Dynamic Resume Generator

![Node.js](https://img.shields.io/badge/Node.js-v14+-green) ![License](https://img.shields.io/badge/License-MIT-blue) ![Status](https://img.shields.io/badge/Status-Active-brightgreen)

A powerful, JSON-driven resume generator that creates professionally formatted **DOCX and PDF** documents from simple JSON data. Perfect for developers who want to version control their resume and easily update content without touching code.

## ✨ Features

- 🎨 **Professional Formatting** - Beautiful, ATS-friendly resume layout with custom colors and styling
- 📝 **JSON-Based Content** - Edit your resume in `resume-data.json` - no code changes needed
- 📄 **Dual Format Output** - Generates both DOCX (Word) and PDF files automatically
- 🔄 **Version Control Friendly** - Keep your resume in Git with easy change tracking
- 🎯 **Fully Customizable** - Modify styling, colors, and layout in JavaScript
- ⚡ **Fast Generation** - Generate resume in seconds
- 🏢 **Enterprise-Ready** - Supports complex work history, skills, projects, awards, and education

## 🚀 Quick Start

### Prerequisites
- Node.js (v14 or higher)
- npm

### Installation

```bash
# Clone the repository
git clone https://github.com/krushnathube/resume.git
cd resume

# Install dependencies
npm install
```

### Usage

#### 1. **Edit Your Content**

Open `resume-data.json` and update your information:

```json
{
  "personal": {
    "name": "YOUR NAME",
    "title": "Your Title",
    "location": "City, Country",
    "phone": "+1 234 567 8900",
    "email": "your.email@example.com",
    "linkedin": "linkedin.com/in/yourprofile/"
  },
  "summary": {
    "content": "Your professional summary..."
  },
  "experience": [
    {
      "company": "Company Name",
      "location": "City",
      "dates": "Jan 2020 - Present",
      "role": "Your Role",
      "bullets": ["Achievement 1", "Achievement 2"]
    }
  ]
  // ... more sections
}
```

#### 2. **Generate Your Resume**

```bash
npm start
```

#### 3. **Output**

The script generates:
- `Krushna_Thube_Resume.docx` - Professional Word document
- `Krushna_Thube_Resume.pdf` - PDF version

Files are saved in `/mnt/user-data/outputs/` directory.

## 📋 JSON Structure

### Sections Available

- **personal** - Name, title, location, contact info, LinkedIn
- **summary** - Professional summary
- **skills** - Technical skills organized by category
- **experience** - Work experience with bullet points
- **projects** - Key projects with descriptions
- **awards** - Awards and recognition
- **certifications** - Certifications and learning paths
- **education** - Degrees and institutions

## 📁 Project Structure

```
.
├── resume.js              # Main generator script
├── resume-data.json       # Your resume content (edit this!)
├── package.json           # Dependencies & scripts
├── README.md              # This file
└── LICENSE                # MIT License
```

## 🎨 Customization

### Modify Styling

Edit these constants in `resume.js`:

```javascript
const NAME_COLOR   = "1F4464";  // Name color (hex)
const BLUE         = "3D71A3";  // Section headers color
const SUBTITLE_BLU = "4A6E8E";  // Subtitle color
const CONTACT_GRAY = "555555";  // Contact info color
// ... more color options
```

### Change Output Filename

Update the paths in `resume.js`:

```javascript
const docxPath = "/path/to/your/Resume.docx";
const pdfPath = "/path/to/your/Resume.pdf";
```

## 📦 Dependencies

- **[docx](https://github.com/dolanmiu/docx)** - Generate DOCX documents
- **[docx-to-pdf](https://github.com/dolanmiu/docx-to-pdf)** - Convert DOCX to PDF

## 🔧 npm Scripts

```bash
npm start          # Generate resume (DOCX + PDF)
npm run generate   # Alias for npm start
```

## 💡 Use Cases

- **Version Control Your Resume** - Track changes over time with Git
- **Bulk Updates** - Update multiple jobs/projects at once in JSON
- **Consistency** - Ensure consistent formatting across all documents
- **Automation** - Integrate with CI/CD pipelines for automatic generation
- **Multiple Formats** - Generate both DOCX and PDF instantly
- **Easy Sharing** - Version control makes it easy to share specific resume versions

## 📋 Example Workflow

```bash
# 1. Clone the repository
git clone https://github.com/krushnathube/resume.git

# 2. Install dependencies
npm install

# 3. Edit resume-data.json with your information

# 4. Generate resume
npm start

# 5. Commit and push
git add .
git commit -m "Update resume"
git push origin main
```

## 🤝 Contributing

Contributions are welcome! Feel free to:
- Report issues
- Suggest improvements
- Submit pull requests
- Add new features

## 📝 License

This project is licensed under the MIT License - see [LICENSE](LICENSE) file for details.

## ⭐ Show Your Support

If you find this project helpful, please give it a star! ⭐

It helps other developers discover this tool and motivates continued development.

## 🙋 Support

- 📧 Open an issue for bug reports or feature requests
- 💬 Discussions welcome for questions and ideas
- 🐛 Found a bug? Please create an issue with details

## 🎯 Roadmap

- [ ] Add more resume templates
- [ ] Support for custom fonts
- [ ] CLI interface for non-developers
- [ ] Online editor
- [ ] Export to more formats (HTML, Markdown)

---

**Created by [Krushna Thube](https://linkedin.com/in/krushna-thube/)** | [GitHub](https://github.com/krushnathube)
