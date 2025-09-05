/**
 * Express.js server for Azure DevOps RTM Generator API
 */

const express = require('express');
const AzureDevOpsRTMGenerator = require('./azure-devops-rtm');
const Database = require('./database');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3001;

// CORS middleware - allow requests from localhost:8080
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', 'http://localhost:8080');
    res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
    res.header('Access-Control-Allow-Credentials', 'true');
    
    // Handle preflight requests
    if (req.method === 'OPTIONS') {
        res.sendStatus(200);
    } else {
        next();
    }
});

// Middleware
app.use(express.json());

// Initialize Database
let database;
let rtmGenerator;

// Initialize the application
async function initializeApp() {
    try {
        // Initialize database
        database = new Database();
        await database.initialize();
        
        // Initialize RTM Generator with default credentials (if available)
        try {
            if (process.env.AZURE_DEVOPS_ORG_URL && process.env.AZURE_DEVOPS_PAT && process.env.AZURE_DEVOPS_PROJECT) {
                rtmGenerator = new AzureDevOpsRTMGenerator(
                    process.env.AZURE_DEVOPS_ORG_URL,
                    process.env.AZURE_DEVOPS_PAT,
                    process.env.AZURE_DEVOPS_PROJECT
                );
                console.log('RTM Generator initialized with environment variables');
            } else {
                console.log('RTM Generator will be initialized per request using stored connections');
            }
        } catch (error) {
            console.log('RTM Generator will be initialized per request using stored connections');
        }
        
    } catch (error) {
        console.error('Failed to initialize application:', error.message);
        process.exit(1);
    }
}

// Helper function to get RTM generator for a connection
async function getRTMGenerator(connectionId) {
    if (!connectionId) {
        throw new Error('Connection ID is required when no default connection is available');
    }
    
    const connection = await database.getConnectionById(connectionId);
    if (!connection) {
        throw new Error(`Connection with ID ${connectionId} not found`);
    }
    
    return new AzureDevOpsRTMGenerator(
        connection.azure_devops_org_url,
        connection.azure_devops_pat,
        connection.azure_devops_project
    );
}

/**
 * POST /connections
 * Create a new Azure DevOps connection
 */
app.post('/connections', async (req, res) => {
    try {
        const { name, azure_devops_org_url, azure_devops_pat, azure_devops_project } = req.body;
        
        // Validate required fields
        if (!name || !azure_devops_org_url || !azure_devops_pat || !azure_devops_project) {
            return res.status(400).json({
                error: 'Bad Request',
                message: 'Missing required fields: name, azure_devops_org_url, azure_devops_pat, azure_devops_project'
            });
        }
        
        // Test the connection before saving
        try {
            const testGenerator = new AzureDevOpsRTMGenerator(
                azure_devops_org_url,
                azure_devops_pat,
                azure_devops_project
            );
            
            await testGenerator.testConnection();
            console.log('Connection test successful');
            
            // Fetch sprints for this connection
            console.log('Fetching sprints for new connection...');
            const sprints = await testGenerator.fetchAllSprints();
            console.log(`Found ${sprints.length} sprints`);
            
            // Save the connection to database
            const connection = await database.addConnection({
                name,
                azure_devops_org_url,
                azure_devops_pat,
                azure_devops_project
            });
            
            // Store sprints in database
            if (sprints.length > 0) {
                const sprintsStored = await database.storeSprints(connection.id, sprints);
                console.log(`Stored ${sprintsStored} sprints for connection ${connection.id}`);
            }
            
            res.status(201).json({
                success: true,
                message: 'Connection created successfully',
                data: {
                    id: connection.id,
                    name: connection.name,
                    azure_devops_org_url: connection.azure_devops_org_url,
                    azure_devops_project: connection.azure_devops_project,
                    created_at: connection.created_at,
                    sprints_count: sprints.length
                }
            });
            
        } catch (connectionError) {
            console.error('Connection test failed:', connectionError.message);
            return res.status(400).json({
                error: 'Connection Test Failed',
                message: `Unable to connect to Azure DevOps: ${connectionError.message}`
            });
        }
        
    } catch (error) {
        console.error('Error creating connection:', error);
        
        if (error.message.includes('already exists')) {
            return res.status(409).json({
                error: 'Conflict',
                message: error.message
            });
        }
        
        res.status(500).json({
            error: 'Internal Server Error',
            message: 'Failed to create connection'
        });
    }
});

/**
 * GET /connections
 * Get all connections (ID and name only)
 */
app.get('/connections', async (req, res) => {
    try {
        const connections = await database.getConnections();
        
        res.json({
            success: true,
            count: connections.length,
            data: connections
        });
        
    } catch (error) {
        console.error('Error fetching connections:', error);
        res.status(500).json({
            error: 'Internal Server Error',
            message: 'Failed to fetch connections',
            details: error.message
        });
    }
});

/**
 * GET /rtm-report
 * Generate RTM report based on story_ids or sprint_name
 * 
 * Query Parameters:
 * - story_ids: comma-separated list of user story IDs (e.g., "123,456,789")
 * - sprint_name: name of the sprint (e.g., "Sprint 40")
 * - connection_id: ID of the connection to use (optional if default connection exists)
 */
app.get('/rtm-report', async (req, res) => {
    try {
        const { story_ids, sprint_name, connection_id } = req.query;

        // Validate that at least one parameter is provided
        if (!story_ids && !sprint_name) {
            return res.status(400).json({
                error: 'Bad Request',
                message: 'Either story_ids or sprint_name parameter is required'
            });
        }

        // Get RTM generator for the specified connection
        const generator = await getRTMGenerator(connection_id);

        let userStoryIds = [];

        // Handle story_ids parameter
        if (story_ids) {
            try {
                userStoryIds = story_ids.split(',').map(id => {
                    const parsedId = parseInt(id.trim());
                    if (isNaN(parsedId)) {
                        throw new Error(`Invalid story ID: ${id.trim()}`);
                    }
                    return parsedId;
                });
            } catch (error) {
                return res.status(400).json({
                    error: 'Bad Request',
                    message: `Invalid story_ids format: ${error.message}`
                });
            }
        }

        // Handle sprint_name parameter
        else if (sprint_name) {
            try {
                console.log(`Fetching user stories for sprint: ${sprint_name}`);
                userStoryIds = await generator.fetchUserStoriesBySprint(sprint_name);
                
                if (userStoryIds.length === 0) {
                    return res.status(404).json({
                        error: 'Not Found',
                        message: `No user stories found for sprint: ${sprint_name}`
                    });
                }
            } catch (error) {
                console.error('Error fetching user stories by sprint:', error);
                return res.status(500).json({
                    error: 'Internal Server Error',
                    message: `Failed to fetch user stories for sprint: ${sprint_name}`
                });
            }
        }

        console.log(`Generating RTM for user story IDs: ${userStoryIds.join(', ')}`);

        // Generate RTM data
        const rtmData = await generator.generateRTM(userStoryIds);

        // Create summary statistics
        const totalUserStories = new Set(rtmData.map(row => row['User Story ID'])).size;
        const totalTestCases = rtmData.filter(row => row['Test Case ID'] !== '').length;
        const coveredUserStories = new Set(
            rtmData.filter(row => row['Status'] === 'Covered').map(row => row['User Story ID'])
        ).size;
        const missingTestCases = rtmData.filter(row => row['Status'] === 'Missing').length;
        const coveragePercentage = totalUserStories > 0 ? 
            ((coveredUserStories / totalUserStories) * 100).toFixed(1) : '0.0';

        const summary = {
            totalUserStories,
            totalTestCases,
            coveredUserStories,
            missingTestCases,
            coveragePercentage: `${coveragePercentage}%`
        };

        // Return the RTM data with summary
        res.json({
            success: true,
            summary,
            data: rtmData,
            metadata: {
                generatedAt: new Date().toISOString(),
                userStoryIds: userStoryIds,
                requestType: story_ids ? 'story_ids' : 'sprint_name',
                requestValue: story_ids || sprint_name,
                connectionId: connection_id || 'default'
            }
        });

    } catch (error) {
        console.error('Error generating RTM report:', error);
        res.status(500).json({
            error: 'Internal Server Error',
            message: 'Failed to generate RTM report',
            details: error.message
        });
    }
});

/**
 * GET /rtm-report/download
 * Generate and download RTM report as Excel file
 * 
 * Query Parameters:
 * - story_ids: comma-separated list of user story IDs (e.g., "123,456,789")
 * - sprint_name: name of the sprint (e.g., "Sprint 40")
 * - filename: optional custom filename (without extension)
 * - connection_id: ID of the connection to use (optional if default connection exists)
 * 
 * Note: Only one of story_ids or sprint_name can be used at a time
 */
app.get('/rtm-report/download', async (req, res) => {
    try {
        const { story_ids, sprint_name, filename, connection_id } = req.query;

        // Validate that at least one parameter is provided
        if (!story_ids && !sprint_name) {
            return res.status(400).json({
                error: 'Bad Request',
                message: 'Either story_ids or sprint_name parameter is required'
            });
        }

        let userStoryIds = [];
        let reportIdentifier = '';

        // Handle story_ids parameter
        if (story_ids) {
            try {
                userStoryIds = story_ids.split(',').map(id => {
                    const parsedId = parseInt(id.trim());
                    if (isNaN(parsedId)) {
                        throw new Error(`Invalid story ID: ${id.trim()}`);
                    }
                    return parsedId;
                });
                reportIdentifier = `Stories_${userStoryIds.join('_')}`;
            } catch (error) {
                return res.status(400).json({
                    error: 'Bad Request',
                    message: `Invalid story_ids format: ${error.message}`
                });
            }
        }

        // Handle sprint_name parameter
        else if (sprint_name) {
            try {
                console.log(`Fetching user stories for sprint: ${sprint_name}`);
                userStoryIds = await (await getRTMGenerator(connection_id)).fetchUserStoriesBySprint(sprint_name);
                
                if (userStoryIds.length === 0) {
                    return res.status(404).json({
                        error: 'Not Found',
                        message: `No user stories found for sprint: ${sprint_name}`
                    });
                }
                reportIdentifier = sprint_name.replace(/[^a-zA-Z0-9]/g, '_');
            } catch (error) {
                console.error('Error fetching user stories by sprint:', error);
                return res.status(500).json({
                    error: 'Internal Server Error',
                    message: `Failed to fetch user stories for sprint: ${sprint_name}`
                });
            }
        }

        console.log(`Generating RTM Excel file for user story IDs: ${userStoryIds.join(', ')}`);

        // Generate RTM data
        const rtmData = await (await getRTMGenerator(connection_id)).generateRTM(userStoryIds);

        // Generate filename
        const timestamp = new Date().toISOString().slice(0, 10);
        const excelFilename = filename ? 
            `${filename}_${timestamp}.xlsx` : 
            `RTM_Report_${reportIdentifier}_${timestamp}.xlsx`;

        // Generate Excel file using the existing exportToExcel method
        const filePath = await (await getRTMGenerator(connection_id)).exportToExcel(rtmData, excelFilename);

        // Set headers for file download
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${excelFilename}"`);
        res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

        // Send the file
        res.download(filePath, excelFilename, (err) => {
            if (err) {
                console.error('Error sending file:', err);
                if (!res.headersSent) {
                    res.status(500).json({
                        error: 'Internal Server Error',
                        message: 'Failed to download file'
                    });
                }
            } else {
                console.log(`File downloaded successfully: ${excelFilename}`);
                
                // Optional: Clean up the file after download
                // Uncomment the following lines if you want to delete the file after download
                /*
                const fs = require('fs');
                setTimeout(() => {
                    fs.unlink(filePath, (unlinkErr) => {
                        if (unlinkErr) {
                            console.error('Error deleting temporary file:', unlinkErr);
                        } else {
                            console.log(`Temporary file deleted: ${filePath}`);
                        }
                    });
                }, 1000); // Delete after 1 second
                */
            }
        });

    } catch (error) {
        console.error('Error generating RTM Excel file:', error);
        if (!res.headersSent) {
            res.status(500).json({
                error: 'Internal Server Error',
                message: 'Failed to generate RTM Excel file',
                details: error.message
            });
        }
    }
});

/**
 * GET /sprints
 * Get all available sprints/iterations from database
 * 
 * Returns a list of all sprints with their details stored in the database
 * 
 * Query Parameter:
 * - connection_id: ID of the connection to use (required)
 * - refresh: Set to 'true' to fetch fresh data from Azure DevOps (optional)
 */
app.get('/sprints', async (req, res) => {
    try {
        const { connection_id, refresh } = req.query;
        
        if (!connection_id) {
            return res.status(400).json({
                error: 'Bad Request',
                message: 'connection_id query parameter is required'
            });
        }
        
        // If refresh is requested, fetch fresh data from Azure DevOps
        if (refresh === 'true') {
            try {
                const generator = await getRTMGenerator(connection_id);
                console.log('Refreshing sprints from Azure DevOps...');
                
                const freshSprints = await generator.fetchAllSprints();
                
                // Update database with fresh data
                const sprintsStored = await database.storeSprints(parseInt(connection_id), freshSprints);
                console.log(`Refreshed and stored ${sprintsStored} sprints`);
                
                return res.json({
                    success: true,
                    count: freshSprints.length,
                    data: freshSprints,
                    metadata: {
                        source: 'azure_devops_fresh',
                        fetchedAt: new Date().toISOString(),
                        connectionId: connection_id,
                        refreshed: true
                    }
                });
            } catch (refreshError) {
                console.error('Error refreshing sprints:', refreshError);
                // Fall back to database data if refresh fails
            }
        }
        
        // Get sprints from database
        console.log(`Fetching sprints from database for connection ${connection_id}...`);
        const sprints = await database.getSprintsByConnectionId(parseInt(connection_id));
        
        if (sprints.length === 0) {
            return res.status(404).json({
                error: 'Not Found',
                message: 'No sprints found for this connection. Try using refresh=true to fetch from Azure DevOps.'
            });
        }
        
        // Return the sprints data
        res.json({
            success: true,
            count: sprints.length,
            data: sprints,
            metadata: {
                source: 'database',
                fetchedAt: new Date().toISOString(),
                connectionId: connection_id,
                refreshed: false
            }
        });
        
    } catch (error) {
        console.error('Error fetching sprints:', error);
        res.status(500).json({
            error: 'Internal Server Error',
            message: 'Failed to fetch sprints'
        });
    }
});

/**
 * DELETE /connections/:id
 * Delete a connection
 */
app.delete('/connections/:id', async (req, res) => {
    try {
        const connectionId = parseInt(req.params.id);
        
        if (isNaN(connectionId)) {
            return res.status(400).json({
                error: 'Bad Request',
                message: 'Invalid connection ID'
            });
        }
        
        // Delete associated sprints first
        const sprintsDeleted = await database.deleteSprintsByConnectionId(connectionId);
        console.log(`Deleted ${sprintsDeleted} sprints for connection ${connectionId}`);
        
        // Delete the connection
        const deleted = await database.deleteConnection(connectionId);
        
        if (!deleted) {
            return res.status(404).json({
                error: 'Not Found',
                message: 'Connection not found'
            });
        }
        
        res.json({
            success: true,
            message: 'Connection and associated sprints deleted successfully',
            data: {
                connectionId: connectionId,
                sprintsDeleted: sprintsDeleted
            }
        });
        
    } catch (error) {
        console.error('Error deleting connection:', error);
        res.status(500).json({
            error: 'Internal Server Error',
            message: 'Failed to delete connection'
        });
    }
});

/**
 * GET /rtm-download
 * Generate RTM KPI statistics in JSON format (same as Excel KPI sheet)
 * 
 * Query Parameters:
 * - story_ids: comma-separated list of user story IDs (e.g., "123,456,789")
 * - sprint_name: name of the sprint (e.g., "Sprint 40")
 * - connection_id: ID of the connection to use (optional if default connection exists)
 * 
 * Note: Only one of story_ids or sprint_name can be used at a time
 */
app.get('/rtm-download', async (req, res) => {
    try {
        const { story_ids, sprint_name, connection_id } = req.query;

        // Validate that at least one parameter is provided
        if (!story_ids && !sprint_name) {
            return res.status(400).json({
                error: 'Bad Request',
                message: 'Either story_ids or sprint_name parameter is required'
            });
        }

        let userStoryIds = [];
        let reportIdentifier = '';

        // Handle story_ids parameter
        if (story_ids) {
            try {
                userStoryIds = story_ids.split(',').map(id => {
                    const parsedId = parseInt(id.trim());
                    if (isNaN(parsedId)) {
                        throw new Error(`Invalid story ID: ${id.trim()}`);
                    }
                    return parsedId;
                });
                reportIdentifier = `Stories_${userStoryIds.join('_')}`;
            } catch (error) {
                return res.status(400).json({
                    error: 'Bad Request',
                    message: `Invalid story_ids format: ${error.message}`
                });
            }
        }

        // Handle sprint_name parameter
        else if (sprint_name) {
            try {
                console.log(`Fetching user stories for sprint: ${sprint_name}`);
                userStoryIds = await (await getRTMGenerator(connection_id)).fetchUserStoriesBySprint(sprint_name);
                
                if (userStoryIds.length === 0) {
                    return res.status(404).json({
                        error: 'Not Found',
                        message: `No user stories found for sprint: ${sprint_name}`
                    });
                }
                reportIdentifier = sprint_name.replace(/[^a-zA-Z0-9]/g, '_');
            } catch (error) {
                console.error('Error fetching user stories by sprint:', error);
                return res.status(500).json({
                    error: 'Internal Server Error',
                    message: `Failed to fetch user stories for sprint: ${sprint_name}`
                });
            }
        }

        console.log(`Generating RTM KPI data for user story IDs: ${userStoryIds.join(', ')}`);

        // Generate RTM data
        const rtmData = await (await getRTMGenerator(connection_id)).generateRTM(userStoryIds);

        // Calculate Overall Coverage KPIs
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

        // Calculate Module Wise Coverage
        const uniqueFeatures = [...new Set(rtmData.map(row => row['Feature']).filter(feature => feature))];
        
        const moduleWiseCoverage = uniqueFeatures.map(feature => {
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

            return {
                feature: feature,
                totalUseCases: totalUseCases,
                positiveCovered: positiveCovered,
                negativeCovered: negativeCovered,
                edgeCasesCovered: edgeCasesCovered,
                integrationCovered: integrationCovered,
                totalCovered: totalCoveredForFeature,
                coveragePercentage: coveragePercentageForFeature,
                pass: passForFeature,
                fail: failForFeature
            };
        });

        // Calculate additional insights
        const uncoveredUserStories = totalUserStories - new Set(
            rtmData.filter(row => row['Status'] === 'Covered').map(row => row['User Story ID'])
        ).size;
        
        const testCasesByType = {
            positive: positiveCount,
            negative: negativeCount,
            edgeCases: edgeCasesCount,
            integration: integrationCount
        };

        const executionSummary = {
            total: passCount + failCount,
            pass: passCount,
            fail: failCount,
            passPercentage: (passCount + failCount) > 0 ? Math.round((passCount / (passCount + failCount)) * 100) : 0,
            failPercentage: (passCount + failCount) > 0 ? Math.round((failCount / (passCount + failCount)) * 100) : 0
        };

        // Build response
        const kpiData = {
            success: true,
            reportIdentifier: reportIdentifier,
            generatedAt: new Date().toISOString(),
            metadata: {
                connectionId: connection_id || 'default',
                userStoryIds: userStoryIds,
                sprintName: sprint_name || null,
                totalRtmRows: rtmData.length
            },
            overallCoverage: {
                totalUserStories: totalUserStories,
                totalModules: totalModules,
                totalUseCases: totalTestCases,
                testCasesByType: testCasesByType,
                coverage: {
                    totalCovered: totalCovered,
                    coveragePercentage: coveragePercentage,
                    uncoveredUserStories: uncoveredUserStories
                },
                execution: executionSummary
            },
            moduleWiseCoverage: moduleWiseCoverage,
            summary: {
                topPerformingModules: moduleWiseCoverage
                    .filter(module => module.totalUseCases > 0)
                    .sort((a, b) => b.coveragePercentage - a.coveragePercentage)
                    .slice(0, 5)
                    .map(module => ({
                        feature: module.feature,
                        coveragePercentage: module.coveragePercentage,
                        totalUseCases: module.totalUseCases
                    })),
                lowPerformingModules: moduleWiseCoverage
                    .filter(module => module.totalUseCases > 0)
                    .sort((a, b) => a.coveragePercentage - b.coveragePercentage)
                    .slice(0, 5)
                    .map(module => ({
                        feature: module.feature,
                        coveragePercentage: module.coveragePercentage,
                        totalUseCases: module.totalUseCases
                    })),
                riskAreas: moduleWiseCoverage
                    .filter(module => module.coveragePercentage < 80 && module.totalUseCases > 0)
                    .map(module => ({
                        feature: module.feature,
                        coveragePercentage: module.coveragePercentage,
                        totalUseCases: module.totalUseCases,
                        gap: module.totalUseCases - module.totalCovered
                    }))
            }
        };

        res.json(kpiData);

    } catch (error) {
        console.error('Error generating RTM KPI data:', error);
        res.status(500).json({
            error: 'Internal Server Error',
            message: 'Failed to generate RTM KPI data'
        });
    }
});

// Health check endpoint
app.get('/health', async (req, res) => {
    try {
        const { connection_id } = req.query;
        const generator = await getRTMGenerator(connection_id);
        
        await generator.testConnection();
        res.json({
            status: 'healthy',
            message: 'Azure DevOps connection is working',
            connectionId: connection_id || 'default',
            timestamp: new Date().toISOString()
        });
    } catch (error) {
        res.status(503).json({
            status: 'unhealthy',
            message: 'Azure DevOps connection failed',
            error: error.message,
            timestamp: new Date().toISOString()
        });
    }
});

// Root endpoint with API documentation
app.get('/', (req, res) => {
    res.json({
        name: 'Azure DevOps RTM Generator API',
        version: '1.0.0',
        endpoints: {
            'POST /connections': {
                description: 'Add a new Azure DevOps connection',
                body: {
                    name: 'Connection name',
                    azure_devops_org_url: 'Azure DevOps organization URL',
                    azure_devops_pat: 'Personal Access Token',
                    azure_devops_project: 'Project name'
                }
            },
            'GET /connections': {
                description: 'Get all connections (ID and name only)',
                parameters: 'None'
            },
            'GET /rtm-report': {
                description: 'Generate Requirements Traceability Matrix report (JSON)',
                parameters: {
                    story_ids: 'Comma-separated list of user story IDs (e.g., "123,456,789")',
                    sprint_name: 'Name of the sprint (e.g., "Sprint 40")',
                    connection_id: 'Connection ID to use (optional)'
                },
                note: 'Only one of story_ids or sprint_name can be used at a time'
            },
            'GET /rtm-report/download': {
                description: 'Generate and download Requirements Traceability Matrix report as Excel file',
                parameters: {
                    story_ids: 'Comma-separated list of user story IDs (e.g., "123,456,789")',
                    sprint_name: 'Name of the sprint (e.g., "Sprint 40")',
                    filename: 'Optional custom filename (without extension)',
                    connection_id: 'Connection ID to use (optional)'
                },
                note: 'Only one of story_ids or sprint_name can be used at a time'
            },
            'GET /sprints': {
                description: 'Get all available sprints/iterations',
                parameters: {
                    connection_id: 'Connection ID to use (optional)',
                    refresh: 'Set to "true" to refresh from Azure DevOps (optional)'
                },
                returns: 'List of sprints with id, name, path, startDate, finishDate, and timeFrame'
            },
            'GET /health': {
                description: 'Check API health and Azure DevOps connection status',
                parameters: {
                    connection_id: 'Connection ID to test (optional)'
                }
            },
            'GET /rtm-download': {
                description: 'Generate RTM KPI statistics in JSON format (same data as Excel KPI sheet)',
                parameters: {
                    story_ids: 'Comma-separated list of user story IDs (e.g., "123,456,789")',
                    sprint_name: 'Name of the sprint (e.g., "Sprint 40")',
                    connection_id: 'Connection ID to use (optional)'
                },
                returns: 'JSON object with overall coverage, module-wise coverage, and summary insights',
                note: 'Only one of story_ids or sprint_name can be used at a time'
            }
        },
        examples: {
            'Add connection': 'POST /connections with JSON body',
            'Get connections': '/connections',
            'JSON report by story IDs': '/rtm-report?story_ids=123,456,789&connection_id=1',
            'Excel download by sprint': '/rtm-report/download?sprint_name=Sprint%2040&connection_id=1',
            'KPI data by sprint': '/rtm-download?sprint_name=Sprint%2040&connection_id=1'
        }
    });
});

// Error handling middleware
app.use((error, req, res, next) => {
    console.error('Unhandled error:', error);
    res.status(500).json({
        error: 'Internal Server Error',
        message: 'An unexpected error occurred'
    });
});

// 404 handler
app.use((req, res) => {
    res.status(404).json({
        error: 'Not Found',
        message: `Endpoint ${req.method} ${req.path} not found`
    });
});

// Graceful shutdown
process.on('SIGINT', async () => {
    console.log('Shutting down gracefully...');
    if (database) {
        await database.close();
    }
    process.exit(0);
});

// Initialize and start server
initializeApp().then(() => {
    app.listen(PORT, () => {
        console.log(`Azure DevOps RTM Generator API server running on port ${PORT}`);
        console.log(`API documentation available at: http://localhost:${PORT}/`);
        console.log(`Health check available at: http://localhost:${PORT}/health`);
        console.log(`RTM report endpoint: http://localhost:${PORT}/rtm-report`);
        console.log(`Connections management: http://localhost:${PORT}/connections`);
    });
});

module.exports = app;
