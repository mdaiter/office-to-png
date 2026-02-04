/**
 * Web Worker for heavy document parsing operations.
 * 
 * This worker handles CPU-intensive document parsing and layout computation
 * off the main UI thread, preventing jank during sheet/page switching.
 * 
 * Key optimization: The document is parsed ONCE when loaded, and the parsed
 * state is kept in WASM memory via WorkerDocumentHolder. Subsequent sheet/page
 * requests only do layout computation, not full re-parsing.
 */

// State
let wasm = null;
let wasmInitialized = false;
let currentRequestId = 0;

// Document holder - keeps parsed document in WASM memory
let docHolder = null;

/**
 * Initialize the WASM module.
 */
async function initWasm() {
    if (wasmInitialized) return true;
    
    try {
        // Import the WASM module
        const wasmModule = await import('./pkg/office_to_png_wasm.js');
        await wasmModule.default();
        wasm = wasmModule;
        
        // Create the document holder
        docHolder = new wasm.WorkerDocumentHolder();
        
        wasmInitialized = true;
        console.log('[Worker] WASM initialized, document holder created');
        return true;
    } catch (err) {
        console.error('[Worker] WASM init error:', err);
        throw err;
    }
}

/**
 * Load and parse a document (called once per document).
 */
function loadDocument(docBytes, docType) {
    if (!wasmInitialized || !docHolder) {
        throw new Error('WASM not initialized');
    }
    
    console.log(`[Worker] Loading ${docType} document (${docBytes.length} bytes)...`);
    const startTime = performance.now();
    
    // This parses the document and keeps it in WASM memory
    docHolder.load(docBytes, docType);
    
    const elapsed = performance.now() - startTime;
    console.log(`[Worker] Document parsed in ${elapsed.toFixed(1)}ms`);
}

/**
 * Clear loaded document from memory.
 */
function clearDocument() {
    if (docHolder) {
        docHolder.clear();
    }
}

/**
 * Parse an XLSX sheet using the already-loaded document.
 */
function parseXlsxSheet(sheetIndex, requestId, canvasWidth, canvasHeight) {
    if (!wasmInitialized || !docHolder) {
        throw new Error('WASM not initialized');
    }
    if (!docHolder.is_loaded()) {
        throw new Error('No document loaded in worker');
    }
    
    // Check for cancellation before starting
    if (requestId !== currentRequestId) {
        console.log('[Worker] Request already cancelled:', requestId);
        return null;
    }
    
    console.log(`[Worker] Parsing sheet ${sheetIndex} (requestId=${requestId})...`);
    const startTime = performance.now();
    
    try {
        // Use the holder's method - doc already parsed, just compute layout
        const buffers = docHolder.parse_xlsx_sheet(
            sheetIndex,
            requestId,
            canvasWidth || 1200,
            canvasHeight || 800
        );
        
        const elapsed = performance.now() - startTime;
        console.log(`[Worker] Sheet ${sheetIndex} layout computed in ${elapsed.toFixed(1)}ms`);
        
        // Check for cancellation after parsing
        if (requestId !== currentRequestId) {
            console.log('[Worker] Request cancelled after parse:', requestId);
            return null;
        }
        
        // Convert JS Array to plain array for transfer
        const transferBuffers = [];
        for (let i = 0; i < buffers.length; i++) {
            transferBuffers.push(buffers[i]);
        }
        
        return transferBuffers;
    } catch (err) {
        console.error('[Worker] Parse XLSX sheet error:', err);
        throw err;
    }
}

/**
 * Parse a DOCX page using the already-loaded document.
 */
function parseDocxPage(pageIndex, requestId) {
    if (!wasmInitialized || !docHolder) {
        throw new Error('WASM not initialized');
    }
    if (!docHolder.is_loaded()) {
        throw new Error('No document loaded in worker');
    }
    
    // Check for cancellation before starting
    if (requestId !== currentRequestId) {
        console.log('[Worker] Request already cancelled:', requestId);
        return null;
    }
    
    console.log(`[Worker] Parsing page ${pageIndex} (requestId=${requestId})...`);
    const startTime = performance.now();
    
    try {
        // Use the holder's method - doc already parsed, just compute layout
        const buffers = docHolder.parse_docx_page(pageIndex, requestId);
        
        const elapsed = performance.now() - startTime;
        console.log(`[Worker] Page ${pageIndex} layout computed in ${elapsed.toFixed(1)}ms`);
        
        // Check for cancellation after parsing
        if (requestId !== currentRequestId) {
            console.log('[Worker] Request cancelled after parse:', requestId);
            return null;
        }
        
        // Convert JS Array to plain array for transfer
        const transferBuffers = [];
        for (let i = 0; i < buffers.length; i++) {
            transferBuffers.push(buffers[i]);
        }
        
        return transferBuffers;
    } catch (err) {
        console.error('[Worker] Parse DOCX page error:', err);
        throw err;
    }
}

/**
 * Handle incoming messages from main thread.
 */
self.onmessage = async function(e) {
    const { type, ...data } = e.data;
    
    try {
        switch (type) {
            case 'init': {
                // Initialize WASM module
                await initWasm();
                self.postMessage({ 
                    type: 'init_complete',
                    success: true 
                });
                break;
            }
            
            case 'load_document': {
                // Load document bytes and parse (sent once per document)
                const { docBytes, docType } = data;
                loadDocument(new Uint8Array(docBytes), docType);
                self.postMessage({
                    type: 'document_loaded',
                    success: true,
                    sheetCount: docHolder.sheet_count(),
                    pageCount: docHolder.page_count()
                });
                break;
            }
            
            case 'clear_document': {
                clearDocument();
                self.postMessage({
                    type: 'document_cleared'
                });
                break;
            }
            
            case 'parse_xlsx_sheet': {
                const { sheetIndex, requestId, canvasWidth, canvasHeight } = data;
                
                // Update current request ID (cancels any pending request)
                currentRequestId = requestId;
                
                const buffers = parseXlsxSheet(sheetIndex, requestId, canvasWidth, canvasHeight);
                
                if (buffers) {
                    // Post with Transferable for zero-copy
                    self.postMessage({
                        type: 'xlsx_sheet_parsed',
                        sheetIndex,
                        requestId,
                        buffers
                    }, buffers); // Transfer ownership of ArrayBuffers
                }
                // If buffers is null, the request was cancelled - don't send anything
                break;
            }
            
            case 'parse_docx_page': {
                const { pageIndex, requestId } = data;
                
                // Update current request ID (cancels any pending request)
                currentRequestId = requestId;
                
                const buffers = parseDocxPage(pageIndex, requestId);
                
                if (buffers) {
                    // Post with Transferable for zero-copy
                    self.postMessage({
                        type: 'docx_page_parsed',
                        pageIndex,
                        requestId,
                        buffers
                    }, buffers); // Transfer ownership of ArrayBuffers
                }
                break;
            }
            
            case 'cancel': {
                // Cancel current request by invalidating its ID
                const { requestId } = data;
                if (currentRequestId === requestId) {
                    currentRequestId = 0;
                }
                // Don't send a response - just silently cancel
                break;
            }
            
            default:
                console.warn('[Worker] Unknown message type:', type);
                self.postMessage({
                    type: 'error',
                    error: `Unknown message type: ${type}`
                });
        }
    } catch (err) {
        console.error('[Worker] Error handling message:', err);
        self.postMessage({
            type: 'error',
            error: err.message || String(err),
            requestId: data.requestId
        });
    }
};

// Signal that worker is ready
self.postMessage({ type: 'ready' });
