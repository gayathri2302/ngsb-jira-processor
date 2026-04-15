# NGSB Jira Processor

A modern React application for processing and visualizing Jira sprint data exports. Built for IQVIA's NextGen Speaker Bureau (NGSB) project to streamline sprint planning and ticket tracking.

## Features

- 📊 **Interactive Pivot Tables** - View tickets by status or assignee with drill-down capabilities
- 📈 **Excel Export** - Generate comprehensive Excel reports with multiple sheets and pivot tables
- 🎨 **Beautiful UI** - Dark-themed interface with custom status color coding
- 🔍 **Advanced Filtering** - Filter by status, epic, assignee, or search across tickets
- 📱 **Responsive Design** - Works seamlessly on desktop and tablet devices

## Getting Started

### Prerequisites

- Node.js 18+
- npm or yarn

### Installation

```bash
# Install dependencies
npm install

# Start development server
npm run dev

# Build for production
npm run build

# Preview production build
npm run preview
```

## Usage

1. **Export from Jira**: Export your Jira tickets as HTML or Excel (.xls/.xlsx)
2. **Upload**: Drag and drop the file into the application or click to browse
3. **Analyze**: View pivot tables, filter tickets, and analyze sprint progress
4. **Export**: Generate a formatted Excel report with multiple sheets

## Project Structure

```
ngsb-jira-processor/
├── src/
│   ├── App.jsx          # Main application component
│   └── main.jsx         # Application entry point
├── public/              # Static assets
├── index.html           # HTML template
├── package.json         # Dependencies and scripts
└── vite.config.js       # Vite configuration
```

## Features in Detail

### Pivot Tables

- **By Status**: View ticket distribution across different statuses
- **By Owner**: See workload distribution across team members
- Click on any cell to see detailed ticket information

### Excel Export

The exported Excel file includes:

- **Summary Sheet**: Overview of all epics with ticket counts
- **Epic Sheets**: Individual sheets for each epic, organized by status
- **Pivot - Status**: Status distribution across epics
- **Pivot - Owners**: Workload distribution across assignees

### Status Tracking

Supports all NGSB workflow statuses:

- Backlog
- To Do
- IN DEVELOPMENT
- Code Review
- Ready For Deployment
- Deployed - Ready for Testing
- QA In Test
- QA Defect
- QA Complete
- Passed QA
- Ready for UAT
- Regression Test
- Dev On Hold

## Technologies Used

- **React 19** - UI framework
- **Vite** - Build tool and dev server
- **SheetJS (xlsx)** - Excel file generation
- **DM Mono & Syne** - Custom fonts for enhanced typography

## Security

This application runs entirely in the browser. No data is sent to any server. All file processing happens locally on your machine.

## License

ISC

## Author

Built for IQVIA NextGen Speaker Bureau team
