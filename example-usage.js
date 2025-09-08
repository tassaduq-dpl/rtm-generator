/**
 * Example usage of Azure DevOps RTM Generator
 */

const { AzureDevOpsRTMGenerator } = require('./azure-devops-rtm');

async function basicExample(...userStoryIds) {
    console.log('Running basic RTM generation example...');
    
    const rtmGenerator = new AzureDevOpsRTMGenerator(
        process.env.AZURE_DEVOPS_ORG_URL,
        process.env.AZURE_DEVOPS_PAT,
        process.env.AZURE_DEVOPS_PROJECT,
    );

    const userStoryIds = await rtmGenerator.getUserStoriesByDateRange(
        weekStart,
        weekEnd,
    );
    
    try {
        // Generate RTM for specific user stories
        const rtmData = await rtmGenerator.generateRTM(userStoryIds);

        const timestamp = new Date().toISOString().slice(0, 10);

        // Export to Excel
        const excelFile = await rtmGenerator.exportToExcel(rtmData, `RTM_${timestamp}.xlsx`);
        
        console.log(`RTM exported to: ${excelFile}`);
        console.log(`Total entries: ${rtmData.length}`);
        
    } catch (error) {
        console.error('Error:', error.message);
    }
}

// Run examples
async function runExamples() {
    try {
        await basicExample(16447, 16446, 16445, 16444, 16443);
        console.log('\n' + '='.repeat(50) + '\n');

        const userStoryIds = await rtmGenerator.fetchUserStoriesBySprint("Sprint 40");
        await basicExample(...userStoryIds);

    } catch (error) {
        console.error('Example execution failed:', error.message);
    }
}

// Run if this file is executed directly
if (require.main === module) {
    runExamples();
}

module.exports = { basicExample };
