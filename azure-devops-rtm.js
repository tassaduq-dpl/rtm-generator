/**
 * Azure DevOps Requirements Traceability Matrix (RTM) Generator
 * 
 * This script fetches user stories and test cases from Azure DevOps and creates
 * a Requirements Traceability Matrix in Excel format.
 * 
 * Requirements:
 * - axios (for HTTP requests)
 * - exceljs (for Excel file generation)
 * - dotenv (for environment variables)
 * 
 * Install with: npm install axios exceljs dotenv
 */

const axios = require('axios');
const ExcelJS = require('exceljs');
const cheerio = require('cheerio');
require('dotenv').config();

/*
* Executes a function in parallel for each item in a list with a configurable concurrency limit.
* 
* This function creates a pool of worker promises that process items from the list concurrently.
* Each worker continuously processes items until the list is exhausted, ensuring optimal
* resource utilization while preventing overwhelming the system with too many concurrent operations.
*/
const parallel = (list, fn, limit = 5) => {
  const iterator = list.keys();
  const promises = [...Array(limit)].map(async (i) => {
    while ((i = iterator.next()).value !== undefined) {
      const item = list[i.value];
      if (item) await fn(item, i.value);
    }
  });
  return Promise.all(promises);
};

class AzureDevOpsRTMGenerator {
    /**
     * Initialize the Azure DevOps RTM Generator
     * 
     * @param {string} organizationUrl - Azure DevOps organization URL (e.g., https://dev.azure.com/yourorg)
     * @param {string} personalAccessToken - Personal Access Token for authentication
     * @param {string} projectName - Name of the Azure DevOps project
     */
    constructor(organizationUrl, personalAccessToken, projectName) {
        this.organizationUrl = organizationUrl.replace(/\/$/, '');
        this.projectName = projectName;
        this.pat = personalAccessToken;
        
        // Create base URL for API calls
        this.apiBaseUrl = `${this.organizationUrl}/${projectName}/_apis`;
        
        // Create authentication header
        const authString = `:${personalAccessToken}`;
        const authBase64 = Buffer.from(authString).toString('base64');
        
        this.headers = {
            'Authorization': `Basic ${authBase64}`,
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        };
        
        // Configure axios instance
        this.axiosInstance = axios.create({
            headers: this.headers,
        });
        
        console.log('Azure DevOps RTM Generator initialized');
    }
    
    /**
     * Test the connection to Azure DevOps
     */
    async testConnection() {
        try {
            const url = `${this.apiBaseUrl}/work/teamsettings/iterations?api-version=7.0&$timeframe=current`;
            await this.axiosInstance.get(url);
            console.log('Successfully connected to Azure DevOps');
            return true;
        } catch (error) {
            console.error('Failed to connect to Azure DevOps:', error.message);
            if (error.response) {
                console.error('Response status:', error.response.status);
                console.error('Response data:', error.response.data);
            }
            throw error;
        }
    }

    /**
     * Fetch user story IDs by sprint name using REST API
     * 
     * @param {string} sprintName - The name of the sprint/iteration
     * @returns {Array<number>} Array of user story IDs in the sprint
     */
    async fetchUserStoriesBySprint(sprintName) {
        try {
            console.log(`Fetching user stories for sprint: ${sprintName}`);
            
            // First, get all iterations to find the sprint ID
            const iterationsUrl = `${this.apiBaseUrl}/work/teamsettings/iterations?api-version=7.0`;
            const iterationsData = await this.makeApiRequest(iterationsUrl);
            
            if (!iterationsData || !iterationsData.value) {
                console.error('Could not fetch iterations');
                return [];
            }
            
            // Find the sprint by name
            const sprint = iterationsData.value.find(iteration => 
                iteration.name.toLowerCase() === sprintName.toLowerCase()
            );
            
            if (!sprint) {
                console.error(`Sprint '${sprintName}' not found`);
                console.log('Available sprints:', iterationsData.value.map(i => i.name));
                return [];
            }
            
            console.log(`Found sprint: ${sprint.name} (ID: ${sprint.id})`);
            
            // Query work items using WIQL (Work Item Query Language)
            const wiqlQuery = {
                query: `
                SELECT [System.Id]
                FROM WorkItems
                WHERE [System.WorkItemType] = 'User Story'
                AND [System.IterationPath] UNDER '${sprint.path || sprint.name}'
                ORDER BY [System.Id]`
            };
            
            const wiqlUrl = `${this.apiBaseUrl}/wit/wiql?api-version=7.0`;
            const queryResult = await this.makeApiRequest(wiqlUrl, 'POST', wiqlQuery);
            
            if (!queryResult || !queryResult.workItems) {
                console.log(`No user stories found in sprint '${sprintName}'`);
                return [];
            }
            
            // Extract user story IDs
            const userStoryIds = queryResult.workItems.map(item => item.id);
            
            console.log(`Found ${userStoryIds.length} user stories in sprint '${sprintName}':`, userStoryIds);
            return userStoryIds;
            
        } catch (error) {
            console.error(`Error fetching user stories for sprint '${sprintName}':`, error.message);
            if (error.response) {
                console.error('Response status:', error.response.status);
                console.error('Response data:', error.response.data);
            }
            return [];
        }
    }
    
    /**
     * Make an API request to Azure DevOps
     * 
     * @param {string} url - API endpoint URL
     * @param {string} method - HTTP method (GET, POST, etc.)
     * @param {Object} data - Request data for POST requests
     * @returns {Object|null} Response data or null if error
     */
    async makeApiRequest(url, method = 'GET', data = null) {
        try {
            let response;
            
            if (method.toUpperCase() === 'GET') {
                response = await this.axiosInstance.get(url);
            } else if (method.toUpperCase() === 'POST') {
                response = await this.axiosInstance.post(url, data);
            } else {
                throw new Error(`Unsupported HTTP method: ${method}`);
            }
            
            return response.data;
            
        } catch (error) {
            console.error(`API request failed for ${url}:`, error.message);
            if (error.response) {
                console.error('Response status:', error.response.status);
                console.error('Response data:', error.response.data);
            }
            return null;
        }
    }

    /**
     * Parse acceptance criteria from HTML string
     * 
     * @param {string} htmlString - HTML string containing acceptance criteria
     * @returns {Object} Object containing acceptance criteria with index as key and text as value
     */
    parseAcceptanceCriterias(htmlString) {
        if (!htmlString) return {};
    
        const $ = cheerio.load(htmlString);
        const criteria = {};
        
        $('li').each((index, element) => {
            const text = $(element).text().trim();
            if (text) {
                criteria[index + 1] = text;
            }
        });
        
        return criteria;
    }
    
    /**
     * Fetch a user story by ID using REST API
     * 
     * @param {number} userStoryId - The ID of the user story
     * @returns {Object|null} Dictionary containing user story details or null if not found
     */
    async fetchUserStory(userStoryId) {
        try {
            const url = `${this.apiBaseUrl}/wit/workitems/${userStoryId}?api-version=7.0`;
            const workItemData = await this.makeApiRequest(url);
            
            if (!workItemData) {
                return null;
            }
            
            const fields = workItemData.fields || {};
            
            if (fields['System.WorkItemType'] !== 'User Story') {
                console.warn(`Work item ${userStoryId} is not a User Story`);
                return null;
            }
            
            const areaPath = fields['System.AreaPath'] || '';
            const feature = areaPath ? areaPath.split('\\').pop() : '';
            
            return {
                id: workItemData.id,
                title: fields['System.Title'] || '',
                description: fields['System.Description'] || '',
                state: fields['System.State'] || '',
                priority: fields['Microsoft.VSTS.Common.Priority'] || 'Medium',
                feature: fields['System.Title'] || feature,
                tags: fields['System.Tags'] || '',
                acceptanceCriteria: this.parseAcceptanceCriterias(fields['Microsoft.VSTS.Common.AcceptanceCriteria'])
            };
        } catch (error) {
            console.error(`Error fetching user story ${userStoryId}:`, error.message);
            return null;
        }
    }
    
    /**
     * Fetch test cases related to a user story using REST API
     * 
     * @param {number} userStoryId - The ID of the user story
     * @returns {Array} List of test case objects
     */
    async fetchRelatedTestCases(userStoryId) {
        try {
            const testCases = [];
            
            // Get work item with relations
            const url = `${this.apiBaseUrl}/wit/workitems/${userStoryId}?$expand=Relations&api-version=7.0`;
            const workItemData = await this.makeApiRequest(url);

            if (!workItemData?.relations?.length) return testCases;
            
            for (const relation of workItemData.relations) {
                // Look for test case relations
                const rel = relation.rel || '';
                if (!(rel === 'Microsoft.VSTS.Common.TestedBy' || rel.toLowerCase().includes('test'))) continue;

                // Extract work item ID from URL
                const relationUrl = relation.url || '';
                if (!relationUrl) continue;

                const testCaseId = parseInt(relationUrl.split('/').pop());
                const testCase = await this.fetchTestCase(testCaseId);
                if (!testCase) continue;

                testCases.push(testCase);
            }

            return testCases;
            
        } catch (error) {
            console.error(`Error fetching test cases for user story ${userStoryId}:`, error.message);
            return [];
        }
    }
    
    /**
     * Fetch a test case by ID using REST API
     * 
     * @param {number} testCaseId - The ID of the test case
     * @returns {Object|null} Dictionary containing test case details or null if not found
     */
    async fetchTestCase(testCaseId) {
        try {
            const url = `${this.apiBaseUrl}/wit/workitems/${testCaseId}?api-version=7.0`;
            const workItemData = await this.makeApiRequest(url);
            
            if (!workItemData) {
                return null;
            }
            
            const fields = workItemData.fields || {};
            
            if (fields['System.WorkItemType'] !== 'Test Case') {
                return null;
            }
            
            return {
                id: workItemData.id,
                title: fields['System.Title'] || '',
                description: fields['System.Description'] || '',
                state: fields['System.State'] || '',
                priority: fields['Microsoft.VSTS.Common.Priority'] || 'Medium',
                executionStatus: ['Done', 'Completed'].includes(fields['System.State']) ? 'Pass' : 'Fail',
                steps: fields['Microsoft.VSTS.TCM.Steps'] || '',
                acceptanceCriteria: fields['Custom.AcceptanceCriteriaNumber'] || '',
                scenarioType: fields['Custom.ScenarioType'] || '',
            };
        } catch (error) {
            console.error(`Error fetching test case ${testCaseId}:`, error.message);
            return null;
        }
    }
    
    /**
     * Get the latest test execution status for a test case using REST API
     * 
     * @param {number} testCaseId - The ID of the test case
     * @returns {string} Execution status (Pass, Fail, or Blank)
     */
    async getTestExecutionStatus(testCaseId) {
        try {
            // Get test runs for this test case
            const url = `${this.apiBaseUrl}/test/runs?api-version=7.0`;
            const testRunsData = await this.makeApiRequest(url);
            
            if (!testRunsData || !testRunsData.value) {
                return 'Blank';
            }
            
            // Look for the most recent test result for this test case
            for (const run of testRunsData.value) {
                const runId = run.id;
                if (runId) {
                    // Get test results for this run
                    const resultsUrl = `${this.apiBaseUrl}/test/runs/${runId}/results?api-version=7.0`;
                    const resultsData = await this.makeApiRequest(resultsUrl);
                    
                    if (resultsData && resultsData.value) {
                        for (const result of resultsData.value) {
                            const testCase = result.testCase || {};
                            if (testCase.id === testCaseId.toString()) {
                                const outcome = (result.outcome || '').toLowerCase();
                                if (outcome === 'passed') {
                                    return 'Pass';
                                } else if (outcome === 'failed') {
                                    return 'Fail';
                                }
                            }
                        }
                    }
                }
            }
            
            return 'Blank';
            
        } catch (error) {
            console.error(`Error getting execution status for test case ${testCaseId}:`, error.message);
            return 'Blank';
        }
    }
    
    /**
     * Determine the scenario type based on title, description, and tags
     * 
     * @param {string} title - Title of the item
     * @param {string} description - Description of the item
     * @param {string} tags - Tags associated with the item
     * @returns {string} Scenario type (Functional, Integration, UI, API, etc.)
     */
    determineScenarioType(title, description, tags) {
        const content = `${title} ${description} ${tags}`.toLowerCase();
        
        if (['api', 'service', 'endpoint'].some(keyword => content.includes(keyword))) {
            return 'API';
        } else if (['ui', 'interface', 'screen', 'page'].some(keyword => content.includes(keyword))) {
            return 'UI';
        } else if (['integration', 'end-to-end', 'e2e'].some(keyword => content.includes(keyword))) {
            return 'Integration';
        } else if (['performance', 'load', 'stress'].some(keyword => content.includes(keyword))) {
            return 'Performance';
        } else if (['security', 'authentication', 'authorization'].some(keyword => content.includes(keyword))) {
            return 'Security';
        } else {
            return 'Functional';
        }
    }
    
    /**
     * Map Azure DevOps priority to RTM priority format
     * 
     * @param {any} azurePriority - Priority value from Azure DevOps
     * @returns {string} Mapped priority (Low, Medium, High)
     */
    mapPriority(azurePriority) {
        if (azurePriority === null || azurePriority === undefined) {
            return 'Medium';
        }
        
        const priorityStr = azurePriority.toString().toLowerCase();
        
        if (['1', 'critical', 'high'].includes(priorityStr)) {
            return 'High';
        } else if (['2', 'medium', 'normal'].includes(priorityStr)) {
            return 'Medium';
        } else if (['3', '4', 'low'].includes(priorityStr)) {
            return 'Low';
        } else {
            return 'Medium';
        }
    }
    
    /**
     * Generate Requirements Traceability Matrix for given user story IDs
     * 
     * @param {Array<number>} userStoryIds - List of user story IDs to process
     * @returns {Array} Array containing the RTM data
     */
    async generateRTM(userStoryIds) {
        const rtmData = [];

        await parallel(userStoryIds, async (userStoryId) => {

            console.log(`Processing User Story ${userStoryId}`);
            
            // Fetch user story
            const userStory = await this.fetchUserStory(userStoryId);
            if (!userStory) {
                console.warn(`Could not fetch user story ${userStoryId}`);
                return;
            }
            
            // Fetch related test cases
            const allTestCases = await this.fetchRelatedTestCases(userStoryId);
            
            if (!allTestCases.length) {
                // Add entry for user story without test cases
                rtmData.push({
                    'User Story ID': userStoryId,
                    'Feature': userStory.feature,
                    'Scenario Type': this.determineScenarioType(
                        userStory.title,
                        userStory.description,
                        userStory.tags
                    ),
                    'Description': userStory.title,
                    'Test Case ID': '',
                    'Status': 'Missing',
                    'Priority': this.mapPriority(userStory.priority),
                    'Execution': ''
                });
                return;
            } 
            const testCasesMap = allTestCases.reduce((acc, testCase) => {
                const acNo = testCase.acceptanceCriteria;
                acc.set(acNo, (acc.get(acNo) || []).concat(testCase));
                return acc;
            }, new Map());
    
            Object.entries(userStory.acceptanceCriteria).forEach(([acNo, criteria]) => {
                const testCases = testCasesMap.get(Number(acNo)) || [];
                const body = {
                    'User Story ID': userStoryId,
                    'Feature': userStory.title,
                }
                if (!testCases.length) {
                    rtmData.push({
                        ...body,
                        'Scenario Type': 'Functional',
                        'Description': criteria,
                        'Test Case ID': '',
                        'Status': 'Missing',
                        'Priority': 'Medium',
                        'Execution': ''
                    });
                    return;
                }
                testCases.forEach((testCase) => {
                    rtmData.push({
                        ...body,
                        'Scenario Type': testCase.scenarioType || 'Functional',
                        'Description': testCase.description || testCase.title,
                        'Test Case ID': testCase.id,
                        'Status': 'Covered',
                        'Priority': this.mapPriority(testCase.priority),
                        'Execution': testCase.executionStatus
                    });
                });
            })
        })
        
        return rtmData;
    }
    
    /**
     * Export RTM data to Excel file
     * 
     * @param {Array} rtmData - Array containing RTM data
     * @param {string} filename - Output filename (optional)
     * @returns {string} Path to the created Excel file
     */
    async exportToExcel(rtmData, filename = null) {
        if (!filename) {
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
            filename = `RTM_Report_${timestamp}.xlsx`;
        }
        
        try {
            const workbook = new ExcelJS.Workbook();
            
            // Create RTM worksheet
            const rtmWorksheet = workbook.addWorksheet('RTM');
            
            // Define columns
            const columns = [
                { header: 'User Story ID', key: 'User Story ID', width: 15 },
                { header: 'Feature', key: 'Feature', width: 20 },
                { header: 'Scenario Type', key: 'Scenario Type', width: 15 },
                { header: 'Description', key: 'Description', width: 50 },
                { header: 'Test Case ID', key: 'Test Case ID', width: 15 },
                { header: 'Status', key: 'Status', width: 12 },
                { header: 'Priority', key: 'Priority', width: 12 },
                { header: 'Execution', key: 'Execution', width: 12 }
            ];
            
            rtmWorksheet.columns = columns;
            
            // Add data
            rtmWorksheet.addRows(rtmData);
            
            // Style the header row
            const headerRow = rtmWorksheet.getRow(1);
            headerRow.font = { bold: true };
            headerRow.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE0E0E0' }
            };
            
            // Add borders to all cells
            rtmWorksheet.eachRow((row, rowNumber) => {
                row.eachCell((cell) => {
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            });
            
            // Create KPI worksheet instead of Summary
            const kpiWorksheet = workbook.addWorksheet('KPI');
            
            // Calculate values for Overall Coverage Section (instead of using complex formulas)
            const totalUserStories = new Set(rtmData.map(row => row['User Story ID'])).size;
            const totalTestCases = rtmData.length;
            const totalModules = new Set(rtmData.map(row => row['Feature']).filter(feature => feature)).size;
            const positiveCount = rtmData.filter(row => row['Scenario Type'] === 'Positive').length;
            const negativeCount = rtmData.filter(row => row['Scenario Type'] === 'Negative').length;
            const edgeCasesCount = rtmData.filter(row => row['Scenario Type'] === 'Edge Case').length;
            const integrationCount = rtmData.filter(row => row['Scenario Type'] === 'Integration').length;
            const totalCovered = rtmData.filter(row => row['Status'] === 'Covered').length;
            const coveragePercentage = totalTestCases > 0 ? Math.round((totalCovered / totalTestCases) * 100) : 0;
            const passCount = rtmData.filter(row => row['Execution'] === 'Pass').length;
            const failCount = rtmData.filter(row => row['Execution'] === 'Fail').length;
            
            // Overall Coverage Section with calculated values
            kpiWorksheet.addRow(['Overall Coverage']);
            kpiWorksheet.addRow(['Total Modules', totalModules]);
            kpiWorksheet.addRow(['Total Uses Cases', totalTestCases]);
            kpiWorksheet.addRow(['Positive', positiveCount]);
            kpiWorksheet.addRow(['Negative', negativeCount]);
            kpiWorksheet.addRow(['Edge Cases', edgeCasesCount]);
            kpiWorksheet.addRow(['Integration', integrationCount]);
            kpiWorksheet.addRow(['Total Covered', totalCovered]);
            kpiWorksheet.addRow(['Coverage %age', coveragePercentage]);
            kpiWorksheet.addRow(['Pass', passCount]);
            kpiWorksheet.addRow(['Fail', failCount]);
            
            // Add empty rows for spacing
            kpiWorksheet.addRow([]);
            kpiWorksheet.addRow([]);
            
            // Module Wise Coverage Section
            const moduleWiseHeaderRow = kpiWorksheet.addRow([
                'Module Wise Coverage',
                '', '', '', '', '', '', '', '', ''
            ]);
            
            const moduleWiseSubHeaderRow = kpiWorksheet.addRow([
                'Feature',
                'Total Use Cases',
                'Positive Covered',
                'Negative Covered',
                'Edge Cases Covered',
                'Integration Cases Covered',
                'Total Covered',
                'Coverage %age',
                'Pass',
                'Fail'
            ]);
            
            // Get unique features from RTM data and calculate their metrics
            const uniqueFeatures = [...new Set(rtmData.map(row => row['Feature']).filter(feature => feature))];
            
            // Add rows for each feature with calculated values
            uniqueFeatures.forEach((feature) => {
                const featureData = rtmData.filter(row => row['Feature'] === feature);
                
                const totalUseCases = featureData.length;
                const positiveCovered = featureData.filter(row => 
                    row['Scenario Type'] === 'Positive' && row['Status'] === 'Covered'
                ).length;
                const negativeCovered = featureData.filter(row => 
                    row['Scenario Type'] === 'Negative' && row['Status'] === 'Covered'
                ).length;
                const edgeCasesCovered = featureData.filter(row => 
                    row['Scenario Type'] === 'Edge Case' && row['Status'] === 'Covered'
                ).length;
                const integrationCovered = featureData.filter(row => 
                    row['Scenario Type'] === 'Integration' && row['Status'] === 'Covered'
                ).length;
                const totalCoveredForFeature = featureData.filter(row => row['Status'] === 'Covered').length;
                const coveragePercentageForFeature = totalUseCases > 0 ? 
                    Math.round((totalCoveredForFeature / totalUseCases) * 100) : 0;
                const passForFeature = featureData.filter(row => row['Execution'] === 'Pass').length;
                const failForFeature = featureData.filter(row => row['Execution'] === 'Fail').length;
                
                kpiWorksheet.addRow([
                    feature,
                    totalUseCases,
                    positiveCovered,
                    negativeCovered,
                    edgeCasesCovered,
                    integrationCovered,
                    totalCoveredForFeature,
                    coveragePercentageForFeature,
                    passForFeature,
                    failForFeature
                ]);
            });
            
            // Style the KPI worksheet
            
            // Style Overall Coverage header
            const overallCoverageHeader = kpiWorksheet.getRow(1);
            overallCoverageHeader.font = { bold: true, size: 14 };
            overallCoverageHeader.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF90EE90' } // Light green
            };
            
            // Style Module Wise Coverage header
            moduleWiseHeaderRow.font = { bold: true, size: 12 };
            moduleWiseHeaderRow.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF90EE90' } // Light green
            };
            
            // Style Module Wise Coverage sub-header
            moduleWiseSubHeaderRow.font = { bold: true };
            moduleWiseSubHeaderRow.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE0E0E0' } // Light gray
            };
            
            // Set column widths for KPI sheet
            kpiWorksheet.columns = [
                { width: 20 }, // Feature
                { width: 15 }, // Total Use Cases
                { width: 15 }, // Positive Covered
                { width: 15 }, // Negative Covered
                { width: 18 }, // Edge Cases Covered
                { width: 20 }, // Integration Cases Covered
                { width: 15 }, // Total Covered
                { width: 15 }, // Coverage %age
                { width: 10 }, // Pass
                { width: 10 }  // Fail
            ];
            
            // Add borders to all cells in KPI sheet
            kpiWorksheet.eachRow((row, rowNumber) => {
                row.eachCell((cell) => {
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                });
            });
            
            // Save the workbook
            await workbook.xlsx.writeFile(filename);
            
            console.log(`RTM exported to ${filename}`);
            return filename;
            
        } catch (error) {
            console.error('Error exporting to Excel:', error.message);
            throw error;
        }
    }

    /**
     * Fetch all available sprints/iterations
     * 
     * @returns {Array} Array of sprint objects with id, name, path, and dates
     */
    async fetchAllSprints() {
        try {
            console.log('Fetching all available sprints...');
            
            // Get all iterations
            const iterationsUrl = `${this.apiBaseUrl}/work/teamsettings/iterations?api-version=7.0`;
            const iterationsData = await this.makeApiRequest(iterationsUrl);
            
            if (!iterationsData || !iterationsData.value) {
                console.error('Could not fetch iterations');
                return [];
            }
            
            // Format sprint data for better usability
            const sprints = iterationsData.value.reduce((list, iteration) => {
                if (iteration.attributes?.timeFrame === 'future') {
                    return list;
                }
                list.push({
                    id: iteration.id,
                    name: iteration.name,
                    path: iteration.path,
                    startDate: iteration.attributes?.startDate || null,
                    finishDate: iteration.attributes?.finishDate || null,
                    timeFrame: iteration.attributes?.timeFrame || null,
                    url: iteration.url
                });
                return list;
            }, []);
            
            console.log(`Found ${sprints.length} sprints`);
            return sprints;
            
        } catch (error) {
            console.error('Error fetching all sprints:', error.message);
            if (error.response) {
                console.error('Response status:', error.response.status);
                console.error('Response data:', error.response.data);
            }
            return [];
        }
    }

    /**
     * Get user stories by date range
     * 
     * @param {Date} startDate - Start date for the range
     * @param {Date} endDate - End date for the range
     * @param {string} projectArea - Optional area path filter
     * @returns {Array<number>} Array of user story IDs
     */
    async getUserStoriesByDateRange(startDate, endDate, projectArea = null) {
        try {
            // Format dates to YYYY-MM-DD format (no time component)
            const startDateStr = startDate.toISOString().split('T')[0];
            const endDateStr = endDate.toISOString().split('T')[0];
            
            // Build WIQL query to get user stories by date range
            let wiqlQuery = `
                SELECT [System.Id]
                FROM WorkItems
                WHERE [System.WorkItemType] = 'User Story'
                AND [System.CreatedDate] >= '${startDateStr}'
                AND [System.CreatedDate] <= '${endDateStr}'
            `;
            
            // Add project area filter if specified
            if (projectArea) {
                wiqlQuery += ` AND [System.AreaPath] UNDER '${projectArea}'`;
            }
            
            wiqlQuery += ` ORDER BY [System.Id]`;
            
            const wiqlUrl = `${this.apiBaseUrl}/wit/wiql?api-version=7.0`;
            const queryResult = await this.makeApiRequest(wiqlUrl, 'POST', { query: wiqlQuery });
            
            if (!queryResult || !queryResult.workItems) {
                return [];
            }
            
            return queryResult.workItems.map(item => item.id);
            
        } catch (error) {
            console.error('Error fetching user stories by date range:', error);
            return [];
        }
    }

    /**
     * Calculate coverage trend over time periods
     * 
     * @param {number} weeks - Number of weeks to analyze (default: 4)
     * @param {string} projectArea - Optional area path filter
     * @returns {Object} Coverage trend data
     */
    async calculateCoverageTrend(weeks = 4, projectArea = null) {
        const trendData = [];

        await parallel(
            [...Array(weeks)].map((_, i) => weeks - i - 1),
            async (i) => {
                const weekEnd = new Date();
                weekEnd.setDate(weekEnd.getDate() - i * 7);

                const weekStart = new Date(weekEnd);
                weekStart.setDate(weekStart.getDate() - 6);

                try {
                    // Get user stories created/updated in this week
                    const userStoryIds = await this.getUserStoriesByDateRange(
                        weekStart,
                        weekEnd,
                        projectArea
                    );

                    if (!userStoryIds.length) {
                        trendData.push({
                            week: `Week ${weeks - i}`,
                            coverage: 0,
                            totalUserStories: 0,
                            coveredUserStories: 0,
                            dateRange: {
                                start: weekStart.toISOString().split("T")[0],
                                end: weekEnd.toISOString().split("T")[0],
                            },
                        });
                        return;
                    }

                    // Generate RTM for these user stories
                    const rtmData = await this.generateRTM(userStoryIds);

                    // Calculate coverage metrics
                    const totalUserStories = new Set(
                        rtmData.map((row) => row["User Story ID"])
                    ).size;
                    const coveredUserStories = new Set(
                        rtmData
                            .filter((row) => row["Status"] === "Covered")
                            .map((row) => row["User Story ID"])
                    ).size;
                    const coveragePercentage =
                        totalUserStories > 0
                            ? Math.round((coveredUserStories / totalUserStories) * 100)
                            : 0;

                    trendData.push({
                        week: `Week ${weeks - i}`,
                        coverage: coveragePercentage,
                        totalUserStories: totalUserStories,
                        coveredUserStories: coveredUserStories,
                        dateRange: {
                            start: weekStart.toISOString().split("T")[0],
                            end: weekEnd.toISOString().split("T")[0],
                        },
                    });
                } catch (weekError) {
                    console.error(
                        `Error calculating coverage for week ${weeks - i}:`,
                        weekError
                    );
                    trendData.push({
                        week: `Week ${weeks - i}`,
                        coverage: 0,
                        totalUserStories: 0,
                        coveredUserStories: 0,
                        error: "Failed to calculate coverage for this week",
                        dateRange: {
                            start: weekStart.toISOString().split("T")[0],
                            end: weekEnd.toISOString().split("T")[0],
                        },
                    });
                }
            }
        );

        // Sort trendData by week order to ensure correct sequence
        trendData.sort((a, b) => {
            const weekA = parseInt(a.week.split(' ')[1]);
            const weekB = parseInt(b.week.split(' ')[1]);
            return weekA - weekB;
        });

        return {
            weeks: weeks,
            data: trendData,
            calculatedAt: new Date().toISOString(),
        };
    }
}

/**
 * Main function to run the RTM generator
 */
async function main() {
    // Configuration - Update these values or use environment variables
    const ORGANIZATION_URL = process.env.AZURE_DEVOPS_ORG_URL || "https://dev.azure.com/yourorganization";
    const PERSONAL_ACCESS_TOKEN = process.env.AZURE_DEVOPS_PAT || "your_pat_token_here";
    const PROJECT_NAME = process.env.AZURE_DEVOPS_PROJECT || "YourProjectName";
    
    // User Story IDs to process - Update this list
    const USER_STORY_IDS = process.env.USER_STORY_IDS ? 
        process.env.USER_STORY_IDS.split(',').map(id => parseInt(id.trim())) :
        [12345, 12346, 12347]; // Update with actual user story IDs
    
    if (PERSONAL_ACCESS_TOKEN === "your_pat_token_here") {
        console.log("Please update the configuration with your Azure DevOps details");
        console.log("You can either:");
        console.log("1. Edit the script and update the variables");
        console.log("2. Set environment variables:");
        console.log("   - AZURE_DEVOPS_ORG_URL");
        console.log("   - AZURE_DEVOPS_PAT");
        console.log("   - AZURE_DEVOPS_PROJECT");
        console.log("   - USER_STORY_IDS (comma-separated)");
        process.exit(1);
    }
    
    try {
        // Initialize the RTM generator
        const rtmGenerator = new AzureDevOpsRTMGenerator(
            ORGANIZATION_URL,
            PERSONAL_ACCESS_TOKEN,
            PROJECT_NAME
        );
        
        // Test connection
        await rtmGenerator.testConnection();
        
        // Generate RTM
        console.log("Starting RTM generation...");
        const rtmData = await rtmGenerator.generateRTM(USER_STORY_IDS);
        
        // Export to Excel
        const excelFile = await rtmGenerator.exportToExcel(rtmData);
        
        console.log(`\nRTM generation completed successfully!`);
        console.log(`Excel file created: ${excelFile}`);
        console.log(`Total rows in RTM: ${rtmData.length}`);
        
        // Display summary
        const uniqueUserStories = new Set(rtmData.map(row => row['User Story ID'])).size;
        const testCasesFound = rtmData.filter(row => row['Test Case ID'] !== '').length;
        const coveredStories = new Set(
            rtmData.filter(row => row['Status'] === 'Covered').map(row => row['User Story ID'])
        ).size;
        
        console.log("\nSummary:");
        console.log(`User Stories processed: ${uniqueUserStories}`);
        console.log(`Test Cases found: ${testCasesFound}`);
        console.log(`Coverage: ${coveredStories}/${uniqueUserStories} user stories`);
        
    } catch (error) {
        console.error('Error in RTM generation:', error.message);
        process.exit(1);
    }
}

// Export the class for use as a module
module.exports = { AzureDevOpsRTMGenerator, parallel };

// Run main function if this file is executed directly
if (require.main === module) {
    main();
}
