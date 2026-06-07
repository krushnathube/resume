# resume
Create dynamic resume

## How to Run

To generate your resume, follow these steps:

### 1. Install Dependencies
```bash
npm install
```

### 2. Edit Your Content (Optional)
All resume content is stored in `resume-data.json`. You can easily edit your resume without touching the JavaScript code:
- Update personal information, skills, experience, projects, awards, certifications, and education
- The JSON file is well-organized and easy to modify
- See [resume-data.json](resume-data.json) for the structure

### 3. Run the Script
```bash
node resume.js
```
or
```bash
npm start
```

### 4. Output
The script will generate both:
- **DOCX Format**: `Krushna_Thube_Resume.docx` - Word document with full formatting
- **PDF Format**: `Krushna_Thube_Resume.pdf` - PDF version of the resume

Files are saved in the `/mnt/user-data/outputs/` directory.

## File Structure
- **resume.js** - Main generator script (dynamically reads from JSON)
- **resume-data.json** - All resume content in JSON format (edit this to update content)
- **package.json** - Dependencies and scripts

## Editing Your Resume

Edit `resume-data.json` to update:
- **personal**: Name, title, location, contact info, LinkedIn
- **summary**: Professional summary
- **skills**: Technical skills organized by category
- **experience**: Work experience with company, role, and bullet points
- **projects**: Key projects with descriptions
- **awards**: Awards and recognition
- **certifications**: Certifications and learning
- **education**: Degrees and institutions

### Requirements
- Node.js (v14 or higher)
- docx package (automatically installed via npm)
- docx-to-pdf package (automatically installed via npm)
