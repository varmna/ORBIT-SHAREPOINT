// ==UserScript==
// @name         ORBIT Beta → SharePoint Auto-Push
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  Auto-pushes annotations to SharePoint when Submit is clicked
// @match        https://orbit-beta.beta.harmony.a2z.com/*
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      amazon.sharepoint.com
// @connect      phonetool.amazon.com
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    const SP_SITE = 'https://amazon.sharepoint.com/sites/ORBITAnnotations';
    const SP_LIST = 'ORBIT Annotations';
    const DB_NAME = 'OrbitAnnotatorDB';
    const DB_STORE = 'sessions';

    // === STATUS BANNER ===
    function showBanner(msg, color) {
        let b = document.getElementById('sp-banner');
        if (!b) {
            b = document.createElement('div');
            b.id = 'sp-banner';
            b.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:99999;padding:10px 20px;font-family:sans-serif;font-size:14px;font-weight:600;text-align:center;transition:opacity 0.3s;';
            document.body.appendChild(b);
        }
        b.style.background = color;
        b.style.color = color === '#ffc107' ? '#333' : 'white';
        b.textContent = msg;
        b.style.opacity = '1';
        if (color !== '#0078d4') setTimeout(function() { b.style.opacity = '0'; }, 5000);
    }

    // === SHAREPOINT API ===
    function spRequest(url, method, body) {
        return new Promise(function(resolve, reject) {
            GM_xmlhttpRequest({
                method: method || 'GET',
                url: url,
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose'
                },
                data: body || null,
                withCredentials: true,
                onload: function(resp) {
                    if (resp.status >= 200 && resp.status < 300) {
                        try { resolve(JSON.parse(resp.responseText)); }
                        catch(e) { resolve({}); }
                    } else {
                        reject(new Error('HTTP ' + resp.status));
                    }
                },
                onerror: function() { reject(new Error('Network error')); }
            });
        });
    }

    async function getDigest() {
        var d = await spRequest(SP_SITE + '/_api/contextinfo', 'POST');
        return d.d.GetContextWebInformation.FormDigestValue;
    }

    var _entityType = null;
    async function getEntityType() {
        if (_entityType) return _entityType;
        var d = await spRequest(SP_SITE + "/_api/web/lists/getbytitle('" + SP_LIST + "')?$select=ListItemEntityTypeFullName");
        _entityType = d.d.ListItemEntityTypeFullName;
        return _entityType;
    }

    async function addListItem(digest, entityType, data) {
        return new Promise(function(resolve, reject) {
            GM_xmlhttpRequest({
                method: 'POST',
                url: SP_SITE + "/_api/web/lists/getbytitle('" + SP_LIST + "')/items",
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-RequestDigest': digest
                },
                data: JSON.stringify(Object.assign({ '__metadata': { 'type': entityType } }, data)),
                withCredentials: true,
                onload: function(resp) {
                    if (resp.status >= 200 && resp.status < 300) resolve(true);
                    else reject(new Error('HTTP ' + resp.status));
                },
                onerror: function() { reject(new Error('Network error')); }
            });
        });
    }

    // === READ FROM INDEXEDDB ===
    function readDB() {
        return new Promise(function(resolve, reject) {
            var req = indexedDB.open(DB_NAME, 1);
            req.onsuccess = function(e) {
                var db = e.target.result;
                var tx = db.transaction(DB_STORE, 'readonly');
                var store = tx.objectStore(DB_STORE);
                var getReq = store.get('active_session');
                getReq.onsuccess = function() { resolve(getReq.result || null); };
                getReq.onerror = function() { reject(new Error('DB read failed')); };
            };
            req.onerror = function() { reject(new Error('DB open failed')); };
        });
    }

    // === TRACK PUSHED CONVERSATIONS ===
    function getPushed() {
        try { return JSON.parse(localStorage.getItem('orbit_sp_pushed') || '{}'); }
        catch(e) { return {}; }
    }
    function markPushed(convId) {
        var d = getPushed();
        d[convId] = new Date().toISOString();
        localStorage.setItem('orbit_sp_pushed', JSON.stringify(d));
    }

    // === BUILD SHAREPOINT ROW ===
    function buildSPItem(row) {
        // Stringify raw metadata if object
        var rawMsgMeta = row.messageMetadata || '';
        if (typeof rawMsgMeta === 'object') rawMsgMeta = JSON.stringify(rawMsgMeta);
        var rawLlmMeta = row.llmMetadata || '';
        if (typeof rawLlmMeta === 'object') rawLlmMeta = JSON.stringify(rawLlmMeta);
        return {
            // SP internal name → value (display names are renamed separately)
            ConversationId: String(row.conversationId || row.Id || '').substring(0, 255),
            MessageIndex: row.messageIndex != null ? Number(row.messageIndex) : 0,
            MessageId: String(row.messageId || '').substring(0, 255),
            CustomerMessage: String(row.llmGeneratedUserMessage || ''),
            BotMessage: String(row.botMessage || ''),
            Feedback: String(row.feedback || '').substring(0, 255),
            UserIntent: String(row.userIntent || '').substring(0, 255),
            CreatedAt: String(row.createdAt || row.CreatedAt || '').substring(0, 255),
            messageMetadata: String(rawMsgMeta),
            llmMetadata: String(rawLlmMeta),
            LlmIntent: String(row.llmIntent || '').substring(0, 255),
            LlmWorkflow: String(row.llmWorkflow || '').substring(0, 255),
            PrimaryOfflineSubIntent: String(row['llmMetadata.primary_offline_sub_intent'] || '').substring(0, 255),
            IsPillMessage: String(row['messageMetadata.isPillMessage'] || '').substring(0, 255),
            PageSourceUrl: String(row['messageMetadata.pageSourceUrl'] || '').substring(0, 255),
            ResponseSource: String(row['llmMetadata.responseSource'] || '').substring(0, 255),
            WeblabOverrides: String(row['messageMetadata.weblabOverrides'] || '').substring(0, 255),
            originPageType: String(row['messageMetadata.originPageType'] || '').substring(0, 255),
            originSubPageType: String(row['messageMetadata.originSubPageType'] || '').substring(0, 255),
            HvaCategory: String(row.hva_category || '').substring(0, 255),
            InteractionType: String(row.interaction_type || '').substring(0, 255),
            StaticResponseType: String(row.static_response_type || '').substring(0, 255),
            CustomerServiceRouting: String(row.customer_service_routing || '').substring(0, 255),
            ResponseContentAccurate: String(row.response_content_accurate || '').substring(0, 255),
            ExpectedResponse: String(row.expected_response || ''),
            Observations: String(row.observations || ''),
            IsBotFirst: String(row['Is Bot First'] || '').substring(0, 255),
            AnnotatorLogin: String(row.Login || '').substring(0, 255),
            AnnotationTimestamp: String(row.Timestamp || '').substring(0, 255),
            CTType: String(row.CT_Type || '').substring(0, 255)
        };
    }

    // === PARSE METADATA FROM RAW JSON/DynamoDB ===
    function parseMetadata(raw) {
        if (!raw || raw === '{}') return {};
        if (typeof raw === 'object') return raw;
        var s = String(raw).trim();
        // Strip outer quotes
        if (s.startsWith('"') && s.endsWith('"')) s = s.slice(1, -1);
        // Replace doubled quotes
        if (s.includes('""')) s = s.replace(/""/g, '"');
        // Try JSON parse
        try { return JSON.parse(s); } catch(e) {}
        // Try after replacing literal \n
        try { return JSON.parse(s.replace(/\\n/g, '\n').replace(/\\r/g, '')); } catch(e) {}
        // DynamoDB format
        if (s.trimStart().startsWith('[')) {
            try {
                var ds = s.replace(/(\{|,)\s*([A-Za-z_][A-Za-z0-9_]*)\s*:/g, '$1"$2":');
                var arr = JSON.parse(ds);
                var obj = {};
                arr.forEach(function(item) {
                    var m = item.M || item;
                    var k = m.key && m.key.S ? m.key.S : m.key;
                    var v = m.value && m.value.S ? m.value.S : m.value;
                    if (k && v !== undefined) obj[k] = v;
                });
                return obj;
            } catch(e) {}
        }
        // Regex fallback
        obj = {};
        var fields = ['originPageType','originSubPageType','originPageUrl','isPillMessage','pageSourceUrl','responseSource','primary_offline_sub_intent','weblabOverrides'];
        fields.forEach(function(f) {
            var re = new RegExp('"' + f + '"\\s*:\\s*"([^"]*)"');
            var match = s.match(re);
            if (match) obj[f] = match[1];
        });
        return obj;
    }

    // === BUILD ROWS FROM SESSION DATA (matches Excel export format exactly) ===
    function buildConversationRows(session, convId) {
        var conv = null;
        for (var i = 0; i < session.conversations.length; i++) {
            var c = session.conversations[i];
            if (c[0] && (c[0].Id === convId || c[0].conversationId === convId)) {
                conv = c;
                break;
            }
        }
        if (!conv) return [];

        var ann = session.annotations[convId] || {};
        var rows = [];
        var userName = session.userName || 'Anonymous';
        var ts = (session.convTimestamps || {})[convId] || '';
        var ctType = conv[0].CT_Type || conv[0].ct_type || conv[0].CT_type || '';

        // Find first customer message index
        var firstCustIdx = -1;
        for (var fi = 0; fi < conv.length; fi++) {
            if (conv[fi].llmGeneratedUserMessage) { firstCustIdx = fi; break; }
        }

        // Helper: find bot message for a customer
        function findBot(custIdx) {
            if (conv[custIdx] && conv[custIdx].botMessage) return conv[custIdx];
            for (var bi = custIdx + 1; bi < conv.length; bi++) {
                if (conv[bi].botMessage) return conv[bi];
            }
            return null;
        }

        // Helper: build one export row (same format as Harmony's buildExportRow)
        function makeRow(m, botMsg, idx, msgAnn, isBotFirst) {
            var src = botMsg || m;
            // Parse metadata from raw JSON/DynamoDB blobs
            var msgMeta = parseMetadata(src.messageMetadata || src.MessageMetadata || '');
            var llmMeta = parseMetadata(src.llmMetadata || '');
            var pageUrl = src['messageMetadata.pageSourceUrl'] || msgMeta.originPageUrl || msgMeta.pageSourceUrl || '';
            var isPill = src['messageMetadata.isPillMessage'] || msgMeta.isPillMessage || '';
            var respSource = src['llmMetadata.responseSource'] || llmMeta.responseSource || '';
            var posi = src['llmMetadata.primary_offline_sub_intent'] || llmMeta.primary_offline_sub_intent || (msgAnn ? msgAnn.primary_offline_sub_intent : '') || '';
            var originPageType = src['messageMetadata.originPageType'] || msgMeta.originPageType || '';
            var originSubPageType = src['messageMetadata.originSubPageType'] || msgMeta.originSubPageType || '';
            var wlRaw = src['messageMetadata.weblabOverrides'] || msgMeta.weblabOverrides || '';
            var weblabStr = '';
            if (wlRaw) {
                try { weblabStr = typeof wlRaw === 'string' && wlRaw.startsWith('{') ? wlRaw : JSON.stringify(wlRaw); }
                catch(e) { weblabStr = String(wlRaw); }
            }
            return {
                conversationId: convId,
                messageIndex: m.messageIndex !== undefined ? m.messageIndex : (m.message_index !== undefined ? m.message_index : idx),
                messageId: m.messageId || m.message_id || '',
                llmGeneratedUserMessage: m.llmGeneratedUserMessage || '',
                botMessage: (botMsg ? botMsg.botMessage : m.botMessage) || '',
                feedback: m.feedback || (botMsg ? botMsg.feedback : '') || '',
                userIntent: m.userIntent || (botMsg ? botMsg.userIntent : '') || '',
                createdAt: m.createdAt || m.CreatedAt || m.created_at || (botMsg ? (botMsg.createdAt || botMsg.CreatedAt || botMsg.created_at) : '') || '',
                messageMetadata: src.messageMetadata || src.MessageMetadata || '',
                llmMetadata: src.llmMetadata || '',
                llmIntent: (msgAnn ? msgAnn.llmIntent : '') || (botMsg ? botMsg.llmIntent : m.llmIntent) || '',
                llmWorkflow: (msgAnn ? msgAnn.llmWorkflow : '') || (botMsg ? botMsg.llmWorkflow : m.llmWorkflow) || '',
                'llmMetadata.primary_offline_sub_intent': posi,
                'messageMetadata.isPillMessage': isPill,
                'messageMetadata.pageSourceUrl': pageUrl,
                'llmMetadata.responseSource': respSource,
                'messageMetadata.weblabOverrides': weblabStr,
                'messageMetadata.originPageType': originPageType,
                'messageMetadata.originSubPageType': originSubPageType,
                hva_category: msgAnn ? (msgAnn.hva_category || '') : '',
                interaction_type: msgAnn ? (msgAnn.interaction_type || '') : '',
                static_response_type: msgAnn ? (msgAnn.static_response_type || '') : '',
                customer_service_routing: msgAnn ? (msgAnn.customer_service_routing || '') : '',
                response_content_accurate: msgAnn ? (msgAnn.response_content_accurate || '') : '',
                expected_response: msgAnn ? (msgAnn.expected_response || '') : '',
                observations: msgAnn ? (msgAnn.observations || '') : '',
                'Is Bot First': isBotFirst ? 'Yes' : 'No',
                Login: userName,
                Timestamp: ts,
                CT_Type: ctType
            };
        }

        // Build rows in index order (bot greetings + annotated customer messages)
        conv.forEach(function(m, idx) {
            if (m.botMessage && !m.llmGeneratedUserMessage) {
                // Standalone bot message (greeting/static)
                rows.push(makeRow(m, null, idx, null, true));
            } else if (m.llmGeneratedUserMessage) {
                // Customer message - only include if fully annotated
                var msgAnn = ann[idx] || null;
                if (!msgAnn || !msgAnn._complete) return;
                var botMsg = findBot(idx);
                rows.push(makeRow(m, botMsg, idx, msgAnn, false));
            }
        });

        return rows;
    }

    // === WATCH FOR SUCCESSFUL SUBMIT ===
    function hookSubmit() {
        var statusEl = document.getElementById('status-message');
        if (!statusEl) {
            setTimeout(hookSubmit, 1000);
            return;
        }

        console.log('[ORBIT→SP] Watching for submissions...');

        // Watch the status message for "Conversation submitted!"
        var observer = new MutationObserver(function(mutations) {
            mutations.forEach(function(m) {
                var text = statusEl.textContent || '';
                if (text.includes('submitted') || text.includes('Submitted')) {
                    handleSubmission();
                }
            });
        });
        observer.observe(statusEl, { childList: true, characterData: true, subtree: true, attributes: true });

        // Also watch for text content changes
        var textObserver = new MutationObserver(function() {
            var text = statusEl.textContent || '';
            if (text.includes('submitted') || text.includes('Submitted')) {
                handleSubmission();
            }
        });
        textObserver.observe(statusEl, { childList: true, characterData: true, subtree: true });
    }

    var _pushing = false;
    async function handleSubmission() {
        if (_pushing) return;
        _pushing = true;

        try {
            // Check for explicit submit signal from the app
            var signal = null;
            try { signal = JSON.parse(localStorage.getItem('orbit_submit_signal') || 'null'); } catch(e) {}
            if (!signal || !signal.convId || (Date.now() - signal.ts > 10000)) {
                // No valid signal or signal is stale (>10s old) — do NOT push
                _pushing = false;
                return;
            }
            // Clear the signal immediately
            localStorage.removeItem('orbit_submit_signal');
            var targetConvId = signal.convId;
            console.log('[ORBIT→SP] Submit signal received for:', targetConvId);

            // Wait for IndexedDB to be updated
            await new Promise(function(r) { setTimeout(r, 1500); });

            var session = await readDB();
            if (!session || !session.annotations) {
                _pushing = false;
                return;
            }

            var pushed = getPushed();
            if (pushed[targetConvId]) {
                console.log('[ORBIT→SP] Already pushed:', targetConvId);
                _pushing = false;
                return;
            }

            var newConvs = [targetConvId];

            if (newConvs.length === 0) {
                _pushing = false;
                return;
            }

            showBanner('⏳ Pushing to SharePoint...', '#0078d4');
            var digest = await getDigest();
            var entityType = await getEntityType();
            var totalOk = 0, totalFail = 0;
            var customerCount = 0;

            for (var c = 0; c < newConvs.length; c++) {
                var cid = newConvs[c];
                var convFail = 0;
                var rows = buildConversationRows(session, cid);

                for (var i = 0; i < rows.length; i++) {
                    try {
                        await addListItem(digest, entityType, buildSPItem(rows[i]));
                        totalOk++;
                        if (rows[i]['Is Bot First'] !== 'Yes') customerCount++;
                    } catch(e) {
                        console.error('[ORBIT→SP] Failed:', e.message);
                        totalFail++;
                        convFail++;
                    }
                }
                if (convFail === 0) { markPushed(cid); } else { console.warn("[ORBIT→SP] Conv " + cid + " NOT marked (" + convFail + " failed)"); }
            }

            // Write push result for app to read
            localStorage.setItem('orbit_beta_sp_result', JSON.stringify({ convId: targetConvId, success: totalFail === 0, customerCount: customerCount, ts: Date.now() }));

            if (totalFail === 0) {
                showBanner('✅ ' + totalOk + ' rows pushed to SharePoint!', '#10b981');
            } else {
                showBanner('⚠️ ' + totalOk + ' ok, ' + totalFail + ' failed', '#ffc107');
            }
        } catch(e) {
            console.error('[ORBIT→SP] Error:', e);
            showBanner('❌ SharePoint error: ' + e.message, '#dc3545');
        } finally {
            _pushing = false;
        }
    }

    // === LOCK SUBMITTED CONVERSATIONS ===

    // === AUTO-DETECT USERNAME VIA MIDWAY (bypasses CORS) ===
    function detectUsername() {
        GM_xmlhttpRequest({
            method: 'GET',
            url: 'https://phonetool.amazon.com/users/me',
            withCredentials: true,
            onload: function(resp) {
                if (resp.status === 200) {
                    // Extract alias from phonetool redirect or page content
                    var match = resp.finalUrl && resp.finalUrl.match(/\/users\/([a-z0-9]+)/i);
                    if (match && match[1] && match[1] !== 'me') {
                        setDetectedUser(match[1]);
                        return;
                    }
                    // Try parsing from page content
                    var bodyMatch = resp.responseText.match(/\"login\"\s*:\s*\"([a-z0-9]+)\"/i);
                    if (!bodyMatch) bodyMatch = resp.responseText.match(/data-login="([a-z0-9]+)"/i);
                    if (!bodyMatch) bodyMatch = resp.responseText.match(/\/users\/([a-z0-9]+)/);
                    if (bodyMatch && bodyMatch[1] && bodyMatch[1] !== 'me') {
                        setDetectedUser(bodyMatch[1]);
                    }
                }
            }
        });
    }

    function setDetectedUser(alias) {
        console.log('[ORBIT→SP] Detected user:', alias);
        localStorage.setItem('annotator_username', alias);
        // Update Harmony app display if visible
        var display = document.getElementById('user-name-display');
        if (display) display.textContent = alias;
    }

    // === INIT ===
    hookSubmit();
    detectUsername();
    console.log('[ORBIT→SP] Auto-push script loaded');

})();
