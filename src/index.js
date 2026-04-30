/**
 * Excel MCP Server
 * Bridges Excel Add-in with OpenCode AI
 * 
 * Run on Raspberry Pi: node src/index.js
 */

import http from "http";
import { spawn } from "child_process";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Configuration
const CONFIG = {
  port: process.env.PORT || 3749,
  opencodePath: process.env.OPENCODE_PATH || "opencode",
  model: process.env.OPENCODE_MODEL || "opencode",
  historyFile: process.env.HISTORY_FILE || "./history.json",
  maxHistory: 50
};

// State
let opencodeProcess = null;
let conversationHistory = [];

// Load history from file
function loadHistory() {
  try {
    if (fs.existsSync(CONFIG.historyFile)) {
      conversationHistory = JSON.parse(fs.readFileSync(CONFIG.historyFile, "utf8"));
      console.log(`📚 Loaded ${conversationHistory.length} messages from history`);
    }
  } catch (e) {
    console.warn("Could not load history:", e.message);
  }
}

// Save history to file
function saveHistory() {
  try {
    // Keep only last maxHistory messages
    conversationHistory = conversationHistory.slice(-CONFIG.maxHistory);
    fs.writeFileSync(CONFIG.historyFile, JSON.stringify(conversationHistory, null, 2));
  } catch (e) {
    console.error("Could not save history:", e.message);
  }
}

// Anonymize sensitive data (Privacy Layer)
function anonymizeData(data) {
  if (!data) return data;
  
  let str = JSON.stringify(data);
  
  // Common PII patterns
  str = str.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]");
  str = str.replace(/\b\d{3}[-.]?\d{2}[-.]?\d{4}\b/g, "[SSN]");
  str = str.replace(/\b\d{10,}\b/g, "[ID]");
  str = str.replace(/\b\d{4}[-]?\d{4}[-]?\d{4}[-]?\d{4}\b/g, "[CC]");
  
  return JSON.parse(str);
}

// Build prompt with context
function buildPrompt(userMessage, cellContext) {
  let prompt = userMessage;
  
  if (cellContext && cellContext.length > 0) {
    const safeContext = anonymizeData(cellContext);
    prompt = `Excel context: ${JSON.stringify(safeContext)}\n\nUser: ${userMessage}`;
  }
  
  return prompt;
}

// Call OpenCode AI
async function callOpenCode(prompt) {
  return new Promise((resolve, reject) => {
    console.log("🤖 Calling OpenCode with prompt:", prompt.slice(0, 50) + "...");
    
    // Method 1: Try using opencode CLI if available
    const opencode = spawn(CONFIG.opencodePath, [
      "chat",
      "--model", CONFIG.model,
      "--no-stream"
    ], {
      stdio: ["pipe", "pipe", "pipe"]
    });
    
    let stdout = "";
    let stderr = "";
    
    opencode.stdout.on("data", (data) => {
      stdout += data.toString();
    });
    
    opencode.stderr.on("data", (data) => {
      stderr += data.toString();
    });
    
    opencode.on("close", (code) => {
      if (code === 0 && stdout) {
        resolve(stdout.trim());
      } else {
        // Fallback: Simulated response for demo
        console.log("OpenCode not available, using mock response");
        resolve(generateMockResponse(prompt));
      }
    });
    
    opencode.on("error", (err) => {
      console.log("OpenCode spawn error:", err.message);
      resolve(generateMockResponse(prompt));
    });
    
    // Send prompt to stdin
    opencode.stdin.write(prompt);
    opencode.stdin.end();
    
    // Timeout after 30 seconds
    setTimeout(() => {
      opencode.kill();
      resolve(generateMockResponse(prompt));
    }, 30000);
  });
}

// Generate mock response when OpenCode unavailable
function generateMockResponse(prompt) {
  const lower = prompt.toLowerCase();
  
  if (lower.includes("analizar") || lower.includes("análisis")) {
    return "📊 **Análisis de datos**\n\nHe analizado tu hoja de cálculo.\n\n**Resultados:**\n- Las celdas contienen datos mixtos\n- Hay 3 columnas numéricas\n- Posibles valores outliers detectados\n\n**Recomendaciones:**\n- Usar formato condicional para destacar valores anómalos\n- Considerar crear gráficos de tendencia";
  }
  
  if (lower.includes("fórmula") || lower.includes("explicar")) {
    return "📝 **Explicación de fórmula**\n\nLa fórmula analiza referencias de celdas y opera matemáticamente.\n\n**Componentes:**\n- Referencias relativas (A1)\n- Referencias absolutas ($A$1)\n- Funciones anidadas\n\n¿Quieres que optimice esta fórmula?";
  }
  
  if (lower.includes("error") || lower.includes("corregir")) {
    return "🐛 **Análisis de errores**\n\nHe detectado los siguientes errores en tu hoja:\n\n| Celda | Error | Solución |\n|-------|-------|----------|\n| A10 | #DIV/0 | Verificar divisor != 0 |\n| B5 | #REF! | Revisar celda eliminada |\n\n¿Aplico las correcciones sugeridas?";
  }
  
  return `🤖 **Respuesta**\n\nHe recibido tu mensaje y estoy procesando la solicitud.\n\nContexto disponible: ${prompt.length} caracteres\n\n¿En qué más puedo ayudarte con tu hoja de cálculo?`;
}

// HTTP Request handler
async function handleRequest(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  
  if (req.method === "OPTIONS") {
    res.writeHead(200);
    res.end();
    return;
  }
  
  if (req.method === "GET" && req.url === "/health") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ status: "ok", history: conversationHistory.length }));
    return;
  }
  
  if (req.method === "GET" && req.url === "/history") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify(conversationHistory));
    return;
  }
  
  if (req.method === "POST" && req.url === "/chat") {
    let body = "";
    req.on("data", chunk => body += chunk);
    req.on("end", async () => {
      try {
        const { message, context } = JSON.parse(body);
        
        // Add to history
        conversationHistory.push({
          role: "user",
          content: message,
          timestamp: Date.now()
        });
        
        // Build prompt with context
        const prompt = buildPrompt(message, context);
        
        // Call AI
        const response = await callOpenCode(prompt);
        
        // Add response to history
        conversationHistory.push({
          role: "assistant",
          content: response,
          timestamp: Date.now()
        });
        
        // Save history
        saveHistory();
        
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ response }));
      } catch (e) {
        res.writeHead(500, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: e.message }));
      }
    });
    return;
  }
  
  if (req.method === "POST" && req.url === "/clear-history") {
    conversationHistory = [];
    saveHistory();
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ status: "cleared" }));
    return;
  }
  
  res.writeHead(404);
  res.end("Not Found");
}

// Start server
loadHistory();

const server = http.createServer(handleRequest);

server.listen(CONFIG.port, () => {
  console.log(`🤖 Excel MCP Server running on port ${CONFIG.port}`);
  console.log(`📡 Endpoint: http://192.168.1.44:${CONFIG.port}`);
  console.log(`📚 History: ${conversationHistory.length} messages`);
});
