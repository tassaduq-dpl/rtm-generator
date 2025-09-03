# Azure DevOps RTM Generator (Node.js)

A Node.js application that fetches user stories and test cases from Azure DevOps and generates a Requirements Traceability Matrix (RTM) in Excel format.

## Features

- Fetches user stories by ID from Azure DevOps using REST API
- Automatically finds related test cases for each user story
- Generates comprehensive RTM with all required columns
- Exports data to Excel with formatting and summary sheet
- Configurable scenario type detection
- Environment variable support for secure configuration
- Comprehensive error handling and logging

## Prerequisites

1. **Node.js**: Version 14.0.0 or higher
2. **Azure DevOps Access**: Access to an Azure DevOps organization and project
3. **Personal Access Token (PAT)**: Create a PAT with Work Items: Read permissions

## Installation

1. Clone or download the project files
2. Install dependencies:
```bash
npm install
```

3. Configure your Azure DevOps settings:
```bash
cp .env.example .env
# Edit .env with your actual values
```

## Configuration

### Option 1: Environment Variables (Recommended)

Create a `.env` file:
```env
AZURE_DEVOPS_ORG_URL=https://dev.azure.com/yourorganization
AZURE_DEVOPS_PAT=your_personal_access_token
AZURE_DEVOPS_PROJECT=YourProjectName
USER_STORY_IDS=12345,12346,12347
```

### Option 2: Direct Configuration

Edit the variables in `azure-devops-rtm.js`:
```javascript
const ORGANIZATION_URL = "https://dev.azure.com/yourorganization";
const PERSONAL_ACCESS_TOKEN = "your_pat_token_here";
const PROJECT_NAME = "YourProjectName";
const USER_STORY_IDS = [12345, 12346, 12347];
```

## Usage

### Command Line Usage

```bash
# Using npm script
npm start

# Direct node execution
node azure-devops-rtm.js
```

### Programmatic Usage

```javascript
const AzureDevOpsRTMGenerator = require('./azure-devops-rtm');

async function generateRTM() {
    const rtmGenerator = new AzureDevOpsRTMGenerator(
        'https://dev.azure.com/yourorg',
        'your_pat_token',
        'YourProject'
    );
    
    // Test connection
    await rtmGenerator.testConnection();
    
    // Generate RTM
    const rtmData = await rtmGenerator.generateRTM([12345, 12346, 12347]);
    
    // Export to Excel
    const excelFile = await rtmGenerator.exportToExcel(rtmData, 'my_rtm.xlsx');
    
    console.log(`RTM exported to: ${excelFile}`);
}

generateRTM().catch(console.error);
```

## Output

The application generates an Excel file with two sheets:

1. **RTM Sheet**: Complete traceability matrix with columns:
   - User Story ID
   - Feature
   - Scenario Type
   - Description
   - Test Case ID
   - Status (Covered, Missing)
   - Priority (Low, Medium, High)
   - Execution (Pass, Fail, Blank)

2. **Summary Sheet**: High-level metrics including:
   - Total User Stories
   - Total Test Cases
   - Coverage Percentage
   - Missing Test Cases

## API Methods

### Class: AzureDevOpsRTMGenerator

#### Constructor
```javascript
new AzureDevOpsRTMGenerator(organizationUrl, personalAccessToken, projectName)
```

#### Methods

- `testConnection()` - Test connection to Azure DevOps
- `fetchUserStory(userStoryId)` - Fetch a user story by ID
- `fetchRelatedTestCases(userStoryId)` - Get related test cases
- `fetchTestCase(testCaseId)` - Fetch a test case by ID
- `generateRTM(userStoryIds)` - Generate RTM data
- `exportToExcel(rtmData, filename)` - Export to Excel file

## Scenario Type Detection

The application automatically determines scenario types based on keywords:

- **API**: api, service, endpoint
- **UI**: ui, interface, screen, page
- **Integration**: integration, end-to-end, e2e
- **Performance**: performance, load, stress
- **Security**: security, authentication, authorization
- **Functional**: Default for items that don't match other categories

## Error Handling

The application includes comprehensive error handling:

- Connection validation on startup
- Graceful handling of missing work items
- Detailed error logging with HTTP status codes
- Fallback values for missing data

## Dependencies

- **axios**: HTTP client for REST API calls
- **exceljs**: Excel file generation and formatting
- **dotenv**: Environment variable management

## Security Notes

- Never commit PAT tokens to version control
- Use environment variables for sensitive configuration
- Regularly rotate your PAT tokens
- Limit PAT permissions to minimum required scope

## Troubleshooting

### Common Issues

1. **Authentication Error**: Verify your PAT has correct permissions
2. **Project Not Found**: Check project name spelling and access permissions
3. **No Test Cases Found**: Ensure test cases are properly linked to user stories
4. **Excel Export Error**: Check file permissions and disk space

### Debug Mode

Enable detailed logging by modifying the axios timeout or adding debug logs:

```javascript
// Add more detailed logging
console.log('Request URL:', url);
console.log('Request headers:', headers);
```

## License

This project is licensed under the MIT License.
