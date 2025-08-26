const express = require('express');
const cors = require('cors');
const { spawn } = require('child_process');
const path = require('path');
const sqlite3 = require('sqlite3').verbose();
const cron = require('node-cron');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Database connection
const dbPath = path.join(__dirname, 'recruitment_web.db');
const db = new sqlite3.Database(dbPath);

// Helper function to run Python scripts with proper encoding
function runPythonScript(scriptArgs) {
    return new Promise((resolve, reject) => {
        const python = spawn('python', ['web_integration.py', ...scriptArgs], {
            env: { 
                ...process.env, 
                PYTHONIOENCODING: 'utf-8',
                PYTHONLEGACYWINDOWSSTDIO: '1'
            },
            stdio: ['pipe', 'pipe', 'pipe']
        });
        
        let dataString = '';
        let errorString = '';

        python.stdout.on('data', (data) => {
            dataString += data.toString('utf8');
        });

        python.stderr.on('data', (data) => {
            errorString += data.toString('utf8');
        });

        python.on('close', (code) => {
            if (code === 0) {
                try {
                    // Clean the data string in case there are any console logs mixed in
                    const lines = dataString.split('\n');
                    let jsonLine = '';
                    for (const line of lines) {
                        if (line.trim().startsWith('{')) {
                            jsonLine = line.trim();
                            break;
                        }
                    }
                    
                    if (jsonLine) {
                        const result = JSON.parse(jsonLine);
                        resolve(result);
                    } else {
                        // If no JSON found, try parsing the entire string
                        const result = JSON.parse(dataString.trim());
                        resolve(result);
                    }
                } catch (parseError) {
                    console.error('JSON Parse Error:', parseError);
                    console.error('Raw output:', dataString);
                    reject(new Error(`Failed to parse Python output: ${parseError.message}`));
                }
            } else {
                console.error('Python script error:', errorString);
                reject(new Error(`Python script failed with code ${code}: ${errorString}`));
            }
        });

        python.on('error', (error) => {
            console.error('Failed to start Python process:', error);
            reject(new Error(`Failed to start Python process: ${error.message}`));
        });
    });
}

// Serve static files
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.get('/admin', (req, res) => {
    res.sendFile(path.join(__dirname, 'admin.html'));
});

app.get('/client', (req, res) => {
    res.sendFile(path.join(__dirname, 'client.html'));
});

// API Routes

// Get admin settings
app.get('/api/admin/settings', async (req, res) => {
    try {
        const settings = {};
        const rows = await new Promise((resolve, reject) => {
            db.all('SELECT setting_key, setting_value FROM admin_settings', (err, rows) => {
                if (err) reject(err);
                else resolve(rows);
            });
        });

        rows.forEach(row => {
            try {
                // Try to parse JSON values
                if (row.setting_value.startsWith('[') || row.setting_value.startsWith('{')) {
                    settings[row.setting_key] = JSON.parse(row.setting_value);
                } else {
                    settings[row.setting_key] = row.setting_value;
                }
            } catch {
                settings[row.setting_key] = row.setting_value;
            }
        });

        res.json({ success: true, settings });
    } catch (error) {
        console.error('Error fetching settings:', error);
        res.json({ success: false, error: error.message });
    }
});

// Update admin settings
app.put('/api/admin/settings', async (req, res) => {
    try {
        const { settings } = req.body;
        
        for (const [key, value] of Object.entries(settings)) {
            const valueStr = Array.isArray(value) || typeof value === 'object' 
                ? JSON.stringify(value) 
                : String(value);
            
            await new Promise((resolve, reject) => {
                db.run(
                    'INSERT OR REPLACE INTO admin_settings (setting_key, setting_value, updated_at) VALUES (?, ?, ?)',
                    [key, valueStr, new Date().toISOString()],
                    (err) => {
                        if (err) reject(err);
                        else resolve();
                    }
                );
            });
        }

        res.json({ success: true });
    } catch (error) {
        console.error('Error updating settings:', error);
        res.json({ success: false, error: error.message });
    }
});

// Generate report
app.post('/api/admin/generate-report', async (req, res) => {
    try {
        const result = await runPythonScript(['generate_report']);
        res.json(result);
    } catch (error) {
        console.error('Error generating report:', error);
        res.json({ success: false, error: error.message });
    }
});

// Send emails
app.post('/api/admin/send-emails', async (req, res) => {
    try {
        const result = await runPythonScript(['send_emails']);
        res.json(result);
    } catch (error) {
        console.error('Error sending emails:', error);
        res.json({ success: false, error: error.message });
    }
});

// Test email
app.post('/api/admin/test-email', async (req, res) => {
    try {
        // For now, just return success - you can implement actual test email logic
        res.json({ success: true, message: 'Test email functionality not yet implemented' });
    } catch (error) {
        console.error('Error testing email:', error);
        res.json({ success: false, error: error.message });
    }
});

// Get system logs
app.get('/api/admin/logs', async (req, res) => {
    try {
        const limit = req.query.limit || 50;
        const logs = await new Promise((resolve, reject) => {
            db.all(
                'SELECT * FROM system_logs ORDER BY created_at DESC LIMIT ?',
                [limit],
                (err, rows) => {
                    if (err) reject(err);
                    else resolve(rows);
                }
            );
        });

        res.json({ success: true, logs });
    } catch (error) {
        console.error('Error fetching logs:', error);
        res.json({ success: false, error: error.message });
    }
});

// Client chat endpoint
app.post('/api/client/chat', async (req, res) => {
    try {
        const { message } = req.body;
        
        // For now, let's generate a report and return relevant info
        const result = await runPythonScript(['generate_report']);
        
        if (result.success) {
            let response = '';
            const data = result.project_data;
            
            if (message.toLowerCase().includes('status') || message.toLowerCase().includes('current')) {
                response = `Current recruitment status: ${data.overall_completes} out of ${data.total_quota} participants completed (${(data.completion_percentage * 100).toFixed(1)}% complete).`;
            } else if (message.toLowerCase().includes('report') || message.toLowerCase().includes('latest')) {
                response = result.analysis;
            } else if (message.toLowerCase().includes('progress')) {
                response = `We've made good progress! Currently at ${data.overall_completes} completes out of ${data.total_quota} target, which is ${(data.completion_percentage * 100).toFixed(1)}% of our goal.`;
            } else {
                response = `Here's the latest update: ${result.analysis}`;
            }
            
            res.json({
                success: true,
                response: response,
                data: data
            });
        } else {
            res.json({
                success: false,
                response: "I'm having trouble accessing the latest data right now. Please try again in a moment."
            });
        }
    } catch (error) {
        console.error('Error in client chat:', error);
        res.json({
            success: false,
            response: "I'm experiencing some technical difficulties. Please try again later."
        });
    }
});

// Initialize database when server starts
function initializeDatabase() {
    return new Promise((resolve, reject) => {
        // Run the Python script to initialize the database
        runPythonScript(['init_db']).then(() => {
            console.log('Database initialized successfully');
            resolve();
        }).catch((error) => {
            console.log('Database may already be initialized or Python script not found');
            // Continue anyway, as the database might already exist
            resolve();
        });
    });
}

// Start server
async function startServer() {
    try {
        await initializeDatabase();
        
        app.listen(PORT, () => {
            console.log(`
🚀 Recruitment Agent Web Server Started!
📊 Admin Dashboard: http://localhost:${PORT}/admin
💬 Client Portal: http://localhost:${PORT}/client
🌐 Main Page: http://localhost:${PORT}/

Server running on port ${PORT}
            `);
        });
    } catch (error) {
        console.error('Failed to start server:', error);
        process.exit(1);
    }
}

// Handle graceful shutdown
process.on('SIGINT', () => {
    console.log('\nGracefully shutting down...');
    db.close((err) => {
        if (err) {
            console.error('Error closing database:', err.message);
        }
        console.log('Database connection closed.');
        process.exit(0);
    });
});

startServer();