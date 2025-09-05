/**
 * SQLite Database module for storing Azure DevOps connections
 */

const sqlite3 = require('sqlite3').verbose();
const path = require('path');

class Database {
    constructor(dbPath = './connections.db') {
        this.dbPath = dbPath;
        this.db = null;
    }

    /**
     * Initialize the database and create tables if they don't exist
     */
    async initialize() {
        return new Promise((resolve, reject) => {
            this.db = new sqlite3.Database(this.dbPath, (err) => {
                if (err) {
                    console.error('Error opening database:', err.message);
                    reject(err);
                    return;
                }
                console.log('Connected to SQLite database');
                this.createTables()
                    .then(() => resolve())
                    .catch(reject);
            });
        });
    }

    /**
     * Create the connections and sprints tables
     */
    async createTables() {
        return new Promise((resolve, reject) => {
            const createConnectionsTableSQL = `
                CREATE TABLE IF NOT EXISTS connections (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    azure_devops_org_url TEXT NOT NULL,
                    azure_devops_pat TEXT NOT NULL,
                    azure_devops_project TEXT NOT NULL,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            `;

            const createSprintsTableSQL = `
                CREATE TABLE IF NOT EXISTS sprints (
                    id TEXT PRIMARY KEY,
                    connection_id INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    path TEXT,
                    start_date TEXT,
                    finish_date TEXT,
                    time_frame TEXT,
                    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (connection_id) REFERENCES connections (id) ON DELETE CASCADE
                )
            `;

            this.db.run(createConnectionsTableSQL, (err) => {
                if (err) {
                    console.error('Error creating connections table:', err.message);
                    reject(err);
                    return;
                }
                console.log('Connections table created or already exists');
                
                // Create sprints table
                this.db.run(createSprintsTableSQL, (err) => {
                    if (err) {
                        console.error('Error creating sprints table:', err.message);
                        reject(err);
                        return;
                    }
                    console.log('Sprints table created or already exists');
                    resolve();
                });
            });
        });
    }

    /**
     * Add a new connection
     * 
     * @param {Object} connection - Connection details
     * @param {string} connection.name - Connection name
     * @param {string} connection.azure_devops_org_url - Azure DevOps organization URL
     * @param {string} connection.azure_devops_pat - Personal Access Token
     * @param {string} connection.azure_devops_project - Project name
     * @returns {Promise<Object>} Created connection with ID
     */
    async addConnection(connection) {
        return new Promise((resolve, reject) => {
            const { name, azure_devops_org_url, azure_devops_pat, azure_devops_project } = connection;
            
            const insertSQL = `
                INSERT INTO connections (name, azure_devops_org_url, azure_devops_pat, azure_devops_project)
                VALUES (?, ?, ?, ?)
            `;

            this.db.run(insertSQL, [name, azure_devops_org_url, azure_devops_pat, azure_devops_project], function(err) {
                if (err) {
                    if (err.code === 'SQLITE_CONSTRAINT_UNIQUE') {
                        reject(new Error(`Connection with name '${name}' already exists`));
                    } else {
                        reject(err);
                    }
                    return;
                }
                
                resolve({
                    id: this.lastID,
                    name,
                    azure_devops_org_url,
                    azure_devops_project,
                    created_at: new Date().toISOString()
                });
            });
        });
    }

    /**
     * Get all connections (ID and name only)
     * 
     * @returns {Promise<Array>} Array of connections with id and name
     */
    async getConnections() {
        return new Promise((resolve, reject) => {
            const selectSQL = `
                SELECT id, name, created_at, updated_at
                FROM connections
                ORDER BY name ASC
            `;

            this.db.all(selectSQL, [], (err, rows) => {
                if (err) {
                    reject(err);
                    return;
                }
                resolve(rows);
            });
        });
    }

    /**
     * Get a specific connection by ID
     * 
     * @param {number} id - Connection ID
     * @returns {Promise<Object|null>} Connection details or null if not found
     */
    async getConnectionById(id) {
        return new Promise((resolve, reject) => {
            const selectSQL = `
                SELECT id, name, azure_devops_org_url, azure_devops_pat, azure_devops_project, created_at, updated_at
                FROM connections
                WHERE id = ?
            `;

            this.db.get(selectSQL, [id], (err, row) => {
                if (err) {
                    reject(err);
                    return;
                }
                resolve(row || null);
            });
        });
    }

    /**
     * Get a specific connection by name
     * 
     * @param {string} name - Connection name
     * @returns {Promise<Object|null>} Connection details or null if not found
     */
    async getConnectionByName(name) {
        return new Promise((resolve, reject) => {
            const selectSQL = `
                SELECT id, name, azure_devops_org_url, azure_devops_pat, azure_devops_project, created_at, updated_at
                FROM connections
                WHERE name = ?
            `;

            this.db.get(selectSQL, [name], (err, row) => {
                if (err) {
                    reject(err);
                    return;
                }
                resolve(row || null);
            });
        });
    }

    /**
     * Update a connection
     * 
     * @param {number} id - Connection ID
     * @param {Object} updates - Fields to update
     * @returns {Promise<boolean>} True if updated, false if not found
     */
    async updateConnection(id, updates) {
        return new Promise((resolve, reject) => {
            const allowedFields = ['name', 'azure_devops_org_url', 'azure_devops_pat', 'azure_devops_project'];
            const updateFields = [];
            const values = [];

            Object.keys(updates).forEach(key => {
                if (allowedFields.includes(key)) {
                    updateFields.push(`${key} = ?`);
                    values.push(updates[key]);
                }
            });

            if (updateFields.length === 0) {
                resolve(false);
                return;
            }

            values.push(new Date().toISOString()); // updated_at
            values.push(id);

            const updateSQL = `
                UPDATE connections 
                SET ${updateFields.join(', ')}, updated_at = ?
                WHERE id = ?
            `;

            this.db.run(updateSQL, values, function(err) {
                if (err) {
                    reject(err);
                    return;
                }
                resolve(this.changes > 0);
            });
        });
    }

    /**
     * Delete a connection
     * 
     * @param {number} id - Connection ID
     * @returns {Promise<boolean>} True if deleted, false if not found
     */
    async deleteConnection(id) {
        return new Promise((resolve, reject) => {
            const deleteSQL = `DELETE FROM connections WHERE id = ?`;

            this.db.run(deleteSQL, [id], function(err) {
                if (err) {
                    reject(err);
                    return;
                }
                resolve(this.changes > 0);
            });
        });
    }

    /**
     * Store sprints for a connection
     * 
     * @param {number} connectionId - Connection ID
     * @param {Array} sprints - Array of sprint objects
     * @returns {Promise<number>} Number of sprints stored
     */
    async storeSprints(connectionId, sprints) {
        return new Promise((resolve, reject) => {
            // First, delete existing sprints for this connection
            const deleteSQL = `DELETE FROM sprints WHERE connection_id = ?`;
            
            this.db.run(deleteSQL, [connectionId], (err) => {
                if (err) {
                    reject(err);
                    return;
                }
                
                if (sprints.length === 0) {
                    resolve(0);
                    return;
                }
                
                // Insert new sprints
                const insertSQL = `
                    INSERT INTO sprints (id, connection_id, name, path, start_date, finish_date, time_frame)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                `;
                
                let insertedCount = 0;
                let errors = [];
                
                const insertSprint = (index) => {
                    if (index >= sprints.length) {
                        if (errors.length > 0) {
                            reject(new Error(`Failed to insert ${errors.length} sprints: ${errors.join(', ')}`));
                        } else {
                            resolve(insertedCount);
                        }
                        return;
                    }
                    
                    const sprint = sprints[index];
                    this.db.run(insertSQL, [
                        sprint.id,
                        connectionId,
                        sprint.name,
                        sprint.path || null,
                        sprint.startDate || null,
                        sprint.finishDate || null,
                        sprint.timeFrame || null
                    ], function(err) {
                        if (err) {
                            errors.push(`Sprint ${sprint.name}: ${err.message}`);
                        } else {
                            insertedCount++;
                        }
                        insertSprint(index + 1);
                    });
                };
                
                insertSprint(0);
            });
        });
    }

    /**
     * Get sprints for a connection
     * 
     * @param {number} connectionId - Connection ID
     * @returns {Promise<Array>} Array of sprint objects
     */
    async getSprintsByConnectionId(connectionId) {
        return new Promise((resolve, reject) => {
            const selectSQL = `
                SELECT id, name, path, start_date as startDate, finish_date as finishDate, 
                       time_frame as timeFrame, created_at, updated_at
                FROM sprints
                WHERE connection_id = ?
                ORDER BY start_date ASC
            `;

            this.db.all(selectSQL, [connectionId], (err, rows) => {
                if (err) {
                    reject(err);
                    return;
                }
                resolve(rows);
            });
        });
    }

    /**
     * Delete sprints for a connection
     * 
     * @param {number} connectionId - Connection ID
     * @returns {Promise<number>} Number of sprints deleted
     */
    async deleteSprintsByConnectionId(connectionId) {
        return new Promise((resolve, reject) => {
            const deleteSQL = `DELETE FROM sprints WHERE connection_id = ?`;

            this.db.run(deleteSQL, [connectionId], function(err) {
                if (err) {
                    reject(err);
                    return;
                }
                resolve(this.changes);
            });
        });
    }

    /**
     * Close the database connection
     */
    async close() {
        return new Promise((resolve, reject) => {
            if (this.db) {
                this.db.close((err) => {
                    if (err) {
                        reject(err);
                        return;
                    }
                    console.log('Database connection closed');
                    resolve();
                });
            } else {
                resolve();
            }
        });
    }
}

module.exports = Database;
