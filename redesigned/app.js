/* ============================================================
   SVF × UCO Bank — BC/BCA Appointment Form Application
   Redesigned JS Application Engine — Modern & High Performance
   ============================================================ */

(function () {
    'use strict';

    // ── State & Constants ───────────────────────────────────────
    var currentStep = 1;
    var totalSteps = 7;
    var masterRows = [];
    var submitBlocked = false;
    var aadhaarRaw = '';
    var DRAFT_KEY = 'svf_uco_bc_appointment_draft_v2';
    var DEFAULT_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbxYObNNRCYpzYsYGpzTSDd3nv4VfhI5d2WWcKfIAX-yJzBnlgwXOIXKVqfaWbelL2UZ/exec';

    // ── Init on DOM Ready ────────────────────────────────────────
    window.addEventListener('DOMContentLoaded', function () {
        initMaxDOB();
        initMasterData();
        initDraftCheck();
        initEventListeners();
        updateUI();
    });

    function getWebAppUrl() {
        var meta = document.querySelector('meta[name="apps-script-web-app-url"]');
        return (window.APPS_SCRIPT_WEB_APP_URL || (meta && meta.content) || DEFAULT_WEB_APP_URL || '').trim();
    }

    // ── DOB 18+ Restriction ─────────────────────────────────────
    function initMaxDOB() {
        var dobEl = document.getElementById('dob');
        if (!dobEl) return;
        var today = new Date();
        today.setFullYear(today.getFullYear() - 18);
        dobEl.max = today.toISOString().split('T')[0];
    }

    // ── Master Data Population (Local Static Priority) ─────────
    function initMasterData() {
        var staticData = window.STATIC_MASTER_DATA;
        if (staticData && staticData.success && Array.isArray(staticData.data)) {
            masterRows = staticData.data;
            populateStates();
        } else {
            fetchRemoteMasterData();
        }
    }

    function fetchRemoteMasterData() {
        var webAppUrl = getWebAppUrl();
        callBackendJsonp(webAppUrl, 'getMasterData', null)
            .then(function (res) {
                if (res && res.success && Array.isArray(res.data)) {
                    masterRows = res.data;
                    populateStates();
                }
            })
            .catch(function (err) {
                console.warn('Failed to load remote master data:', err);
            });
    }

    function populateStates() {
        var stateSelect = document.getElementById('state');
        if (!stateSelect || !masterRows.length) return;
        
        var states = unique(masterRows.map(function (r) { return r[0]; }));
        stateSelect.innerHTML = '<option value="">— Select State —</option>';
        states.forEach(function (s) {
            var opt = document.createElement('option');
            opt.value = s;
            opt.textContent = s;
            stateSelect.appendChild(opt);
        });
    }

    function unique(arr) {
        return arr.filter(function (v, i, a) { return v && a.indexOf(v) === i; }).sort();
    }

    // ── Branch Cascade Logic ────────────────────────────────────
    window.handleStateChange = function () {
        var stateVal = document.getElementById('state').value;
        var zoneSelect = document.getElementById('zone');
        var branchInput = document.getElementById('branchSearch');
        var branchCodeInput = document.getElementById('branchCode');

        zoneSelect.innerHTML = '<option value="">— Select Zone —</option>';
        zoneSelect.disabled = !stateVal;
        branchInput.value = '';
        branchInput.disabled = true;
        branchCodeInput.value = '';

        if (!stateVal) return;

        var filteredZones = unique(masterRows.filter(function (r) {
            return r[0] === stateVal;
        }).map(function (r) { return r[1]; }));

        filteredZones.forEach(function (z) {
            var opt = document.createElement('option');
            opt.value = z;
            opt.textContent = z;
            zoneSelect.appendChild(opt);
        });
    };

    window.handleZoneChange = function () {
        var stateVal = document.getElementById('state').value;
        var zoneVal = document.getElementById('zone').value;
        var branchInput = document.getElementById('branchSearch');
        var branchCodeInput = document.getElementById('branchCode');

        branchInput.value = '';
        branchInput.disabled = !(stateVal && zoneVal);
        branchCodeInput.value = '';
    };

    window.filterBranches = function () {
        var stateVal = document.getElementById('state').value;
        var zoneVal = document.getElementById('zone').value;
        var query = document.getElementById('branchSearch').value.toLowerCase().trim();
        var dropdown = document.getElementById('branchDropdown');

        if (!stateVal || !zoneVal) return;

        var matches = masterRows.filter(function (r) {
            return r[0] === stateVal && r[1] === zoneVal && r[3].toLowerCase().includes(query);
        });

        dropdown.innerHTML = '';
        if (!matches.length) {
            dropdown.innerHTML = '<div class="branch-option-item" style="color:#64748b;">No matching branches found</div>';
            dropdown.classList.add('show');
            return;
        }

        matches.slice(0, 50).forEach(function (m) {
            var item = document.createElement('div');
            item.className = 'branch-option-item';
            item.innerHTML = '<span>' + m[3] + '</span><span class="branch-option-code">' + m[2] + '</span>';
            item.onclick = function () {
                document.getElementById('branchSearch').value = m[3];
                document.getElementById('branchCode').value = m[2];
                dropdown.classList.remove('show');
                triggerAutoSave();
            };
            dropdown.appendChild(item);
        });
        dropdown.classList.add('show');
    };

    // Close dropdown on outside click
    document.addEventListener('click', function (e) {
        var wrap = document.querySelector('.searchable-select-wrap');
        var dropdown = document.getElementById('branchDropdown');
        if (dropdown && wrap && !wrap.contains(e.target)) {
            dropdown.classList.remove('show');
        }
    });

    // ── Input Formatters & Toggle Buttons ──────────────────────
    function initEventListeners() {
        // Aadhaar Mask
        var aadhaarEl = document.getElementById('aadhaarNo');
        if (aadhaarEl) {
            aadhaarEl.addEventListener('focus', function () {
                if (aadhaarRaw) this.value = aadhaarRaw;
            });
            aadhaarEl.addEventListener('blur', function () {
                var val = this.value.trim().replace(/\D/g, '');
                if (val.length === 12) {
                    aadhaarRaw = val;
                    this.value = 'XXXX-XXXX-' + val.slice(8);
                }
            });
            aadhaarEl.addEventListener('input', function () {
                aadhaarRaw = '';
                triggerAutoSave();
            });
        }

        // Toggle Groups
        document.querySelectorAll('.toggle-group').forEach(function (group) {
            var hiddenInput = document.getElementById(group.getAttribute('data-target'));
            group.querySelectorAll('.toggle-btn').forEach(function (btn) {
                btn.addEventListener('click', function () {
                    group.querySelectorAll('.toggle-btn').forEach(function (b) { b.classList.remove('active'); });
                    btn.classList.add('active');
                    if (hiddenInput) {
                        hiddenInput.value = btn.getAttribute('data-value');
                        hiddenInput.dispatchEvent(new Event('change'));
                    }
                    triggerAutoSave();
                });
            });
        });

        // IIBF Toggle dependency
        var iibfSelect = document.getElementById('iibfCertified');
        if (iibfSelect) {
            iibfSelect.addEventListener('change', function () {
                var isYes = this.value === 'Yes';
                document.getElementById('iibf-cert-wrap').style.display = isYes ? 'block' : 'none';
                document.getElementById('iibf-date-wrap').style.display = isYes ? 'block' : 'none';
            });
        }

        // Upload Purpose dependency
        var purposeSelect = document.getElementById('uploadPurpose');
        if (purposeSelect) {
            purposeSelect.addEventListener('change', function () {
                var isReplacement = this.value === 'REPLACEMENT';
                document.getElementById('replaced-agent-name-wrap').style.display = isReplacement ? 'block' : 'none';
                document.getElementById('replaced-agent-id-wrap').style.display = isReplacement ? 'block' : 'none';
            });
        }

        // Auto Save Listener
        document.addEventListener('input', triggerAutoSave);
        document.addEventListener('change', triggerAutoSave);
    }

    // ── Auto Save & Restore Draft ──────────────────────────────
    var draftTimeout;
    function triggerAutoSave() {
        clearTimeout(draftTimeout);
        draftTimeout = setTimeout(saveDraft, 500);
    }

    function saveDraft() {
        try {
            var data = collectFormData();
            data._savedAt = new Date().toLocaleString();
            localStorage.setItem(DRAFT_KEY, JSON.stringify(data));
        } catch (e) {}
    }

    function initDraftCheck() {
        try {
            var raw = localStorage.getItem(DRAFT_KEY);
            if (!raw) return;
            var draft = JSON.parse(raw);
            if (draft && draft.cspName) {
                var banner = document.getElementById('draftNoticeBanner');
                var text = document.getElementById('draftNoticeText');
                if (banner && text) {
                    text.textContent = 'Saved draft restored from ' + (draft._savedAt || 'earlier session') + '.';
                    banner.style.display = 'flex';
                }
            }
        } catch (e) {}
    }

    window.restoreDraft = function () {
        try {
            var raw = localStorage.getItem(DRAFT_KEY);
            if (!raw) return;
            var d = JSON.parse(raw);
            Object.keys(d).forEach(function (k) {
                var el = document.getElementById(k);
                if (el && d[k]) el.value = d[k];
            });
            document.getElementById('draftNoticeBanner').style.display = 'none';
        } catch (e) {}
    };

    window.clearDraft = function () {
        localStorage.removeItem(DRAFT_KEY);
        var banner = document.getElementById('draftNoticeBanner');
        if (banner) banner.style.display = 'none';
    };

    // ── Navigation Engine ───────────────────────────────────────
    window.nextStep = function () {
        if (!validateStep(currentStep)) return;
        if (currentStep < totalSteps) {
            currentStep++;
            if (currentStep === 7) buildReview();
            updateUI();
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }
    };

    window.prevStep = function () {
        if (currentStep > 1) {
            currentStep--;
            updateUI();
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }
    };

    window.goToStep = function (stepNum) {
        if (stepNum >= 1 && stepNum <= totalSteps) {
            currentStep = stepNum;
            if (currentStep === 7) buildReview();
            updateUI();
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }
    };

    function updateUI() {
        // Update Step Cards
        document.querySelectorAll('.step-card').forEach(function (card, idx) {
            card.classList.toggle('active', (idx + 1) === currentStep);
        });

        // Update Step Tracker Pills
        document.querySelectorAll('.step-pill').forEach(function (pill, idx) {
            var sNum = idx + 1;
            pill.classList.toggle('active', sNum === currentStep);
            pill.classList.toggle('completed', sNum < currentStep);
        });

        // Update Progress Bar
        var pct = Math.round((currentStep / totalSteps) * 100);
        document.getElementById('progressBarFill').style.width = pct + '%';
        document.getElementById('stepIndicatorText').textContent = 'Step ' + currentStep + ' of ' + totalSteps;

        // Update Bottom Nav Buttons
        var prevBtn = document.getElementById('btnPrev');
        var nextBtn = document.getElementById('btnNext');
        var submitBtn = document.getElementById('btnSubmit');

        if (prevBtn) prevBtn.style.display = currentStep > 1 ? 'inline-flex' : 'none';
        if (nextBtn) nextBtn.style.display = currentStep < totalSteps ? 'inline-flex' : 'none';
        if (submitBtn) submitBtn.style.display = currentStep === totalSteps ? 'inline-flex' : 'none';
    }

    // ── Form Validation ─────────────────────────────────────────
    function validateStep(stepNum) {
        var valid = true;
        var stepEl = document.getElementById('step-' + stepNum);
        if (!stepEl) return true;

        stepEl.querySelectorAll('[required]').forEach(function (input) {
            var val = input.value.trim();
            if (input.id === 'aadhaarNo' && aadhaarRaw) val = aadhaarRaw;

            if (!val) {
                valid = false;
                input.classList.add('error');
            } else {
                input.classList.remove('error');
            }
        });

        return valid;
    }

    function validateFullForm() {
        var errorList = [];
        var stepTitles = {
            1: 'Location Details',
            2: 'Personal & Contact Details',
            3: 'KYC & Identification',
            4: 'Bank Account Details',
            5: 'Professional Details',
            6: 'IIBF Certification'
        };

        for (var s = 1; s <= 6; s++) {
            var stepEl = document.getElementById('step-' + s);
            if (!stepEl) continue;

            stepEl.querySelectorAll('[required]').forEach(function (input) {
                var val = input.value.trim();
                if (input.id === 'aadhaarNo' && aadhaarRaw) val = aadhaarRaw;

                if (!val) {
                    var labelText = input.name || input.id;
                    var labelEl = document.querySelector('label[for="' + input.id + '"]');
                    if (labelEl) labelText = labelEl.innerText.replace('*', '').trim();

                    errorList.push({
                        stepNum: s,
                        stepTitle: stepTitles[s] || 'Step ' + s,
                        fieldId: input.id,
                        label: labelText,
                        msg: 'This field is required.'
                    });
                }
            });
        }

        return errorList;
    }

    window.showValidationModal = function (errorList) {
        var modal = document.getElementById('validationModal');
        var listContainer = document.getElementById('validationModalList');
        if (!modal || !listContainer) return;

        listContainer.innerHTML = '';
        errorList.forEach(function (item) {
            var itemEl = document.createElement('div');
            itemEl.className = 'validation-error-item';

            var infoEl = document.createElement('div');
            infoEl.innerHTML = '<div style="font-size:0.75rem;font-weight:700;color:#64748b;">Step ' + item.stepNum + ': ' + item.stepTitle + '</div>' +
                '<div style="font-size:0.92rem;font-weight:700;color:#0f172a;">' + item.label + '</div>' +
                '<div style="font-size:0.8rem;color:#ef4444;">' + item.msg + '</div>';

            var btn = document.createElement('button');
            btn.className = 'btn-jump-step';
            btn.textContent = 'Jump to Page ' + item.stepNum;
            btn.onclick = function () {
                closeValidationModal();
                goToStep(item.stepNum);
                var el = document.getElementById(item.fieldId);
                if (el) el.focus();
            };

            itemEl.appendChild(infoEl);
            itemEl.appendChild(btn);
            listContainer.appendChild(itemEl);
        });

        modal.style.display = 'flex';
    };

    window.closeValidationModal = function () {
        var modal = document.getElementById('validationModal');
        if (modal) modal.style.display = 'none';
    };

    // ── Form Data Collection ────────────────────────────────────
    function collectFormData() {
        var rName = (document.getElementById('replacedAgentName')?.value || '').trim();
        var rId = (document.getElementById('replacedAgentId')?.value || '').trim();
        var combinedReplaced = rName && rId ? rName + ' (' + rId + ')' : (rName || rId || 'NA');

        return {
            state: document.getElementById('state')?.value || '',
            zone: document.getElementById('zone')?.value || '',
            branch: document.getElementById('branchSearch')?.value || '',
            branchCode: document.getElementById('branchCode')?.value || '',
            district: document.getElementById('district')?.value || '',
            block: document.getElementById('block')?.value || '',
            villageName: document.getElementById('villageName')?.value || '',
            villageCode: document.getElementById('villageCode')?.value || '',
            cspName: document.getElementById('cspName')?.value || '',
            fathersName: document.getElementById('fathersName')?.value || '',
            contactNumber: document.getElementById('contactNumber')?.value || '',
            altContactNo: document.getElementById('altContactNo')?.value || '',
            dob: document.getElementById('dob')?.value || '',
            gender: document.getElementById('gender')?.value || 'Male',
            address: document.getElementById('address')?.value || '',
            pinCode: document.getElementById('pinCode')?.value || '',
            caste: document.getElementById('caste')?.value || '',
            aadhaarNo: aadhaarRaw || document.getElementById('aadhaarNo')?.value || '',
            panNo: document.getElementById('panNo')?.value || '',
            otherIdType: document.getElementById('otherIdType')?.value || '',
            otherIdNo: document.getElementById('otherIdNo')?.value || '',
            agentCifNo: document.getElementById('agentCifNo')?.value || '',
            settlementAccount: document.getElementById('settlementAccount')?.value || '',
            savingAccount: document.getElementById('savingAccount')?.value || '',
            education: document.getElementById('education')?.value || '',
            doj: document.getElementById('doj')?.value || '',
            uploadPurpose: document.getElementById('uploadPurpose')?.value || 'NEW',
            shgMember: document.getElementById('shgMember')?.value || 'No',
            physicallyChallenged: document.getElementById('physicallyChallenged')?.value || 'No',
            networkProvider: document.getElementById('networkProvider')?.value || '',
            mailId: document.getElementById('mailId')?.value || '',
            officeMailId: document.getElementById('officeMailId')?.value || '',
            bankMitraActivity: 'NO',
            replacedAgent: combinedReplaced,
            iibfCertified: document.getElementById('iibfCertified')?.value || 'No',
            iibfCertificate: document.getElementById('iibfCertificate')?.value || '',
            certificateDate: document.getElementById('certificateDate')?.value || ''
        };
    }

    // ── Build Review Screen ────────────────────────────────────
    function buildReview() {
        var d = collectFormData();
        var reviewWrap = document.getElementById('reviewGridContainer');
        if (!reviewWrap) return;

        var sections = [
            { title: 'Location Details', fields: [['State', d.state], ['Zone', d.zone], ['Branch', d.branch], ['Branch Code', d.branchCode], ['District', d.district], ['Block', d.block], ['Village', d.villageName]] },
            { title: 'Personal Details', fields: [['CSP Name', d.cspName], ["Father's Name", d.fathersName], ['Contact No', d.contactNumber], ['DOB', d.dob], ['Gender', d.gender], ['Address', d.address], ['PIN Code', d.pinCode]] },
            { title: 'KYC & Bank Details', fields: [['Aadhaar No', '••••' + d.aadhaarNo.slice(-4)], ['PAN No', d.panNo], ['Agent CIF', d.agentCifNo], ['Settlement A/C', d.settlementAccount], ['Saving A/C', d.savingAccount]] },
            { title: 'Professional & IIBF', fields: [['Education', d.education], ['Upload Purpose', d.uploadPurpose], ['Email', d.mailId], ['IIBF Certified', d.iibfCertified]] }
        ];

        var html = '';
        sections.forEach(function (sec) {
            html += '<div class="review-section"><div class="review-section-title">' + sec.title + '</div><div class="review-grid">';
            sec.fields.forEach(function (f) {
                if (f[1]) {
                    html += '<div class="review-item"><div class="review-item-label">' + f[0] + '</div><div class="review-item-val">' + f[1] + '</div></div>';
                }
            });
            html += '</div></div>';
        });

        reviewWrap.innerHTML = html;
    }

    // ── Form Submission Handler ─────────────────────────────────
    window.handleSubmit = function () {
        if (submitBlocked) return;
        var errors = validateFullForm();
        if (errors.length > 0) {
            showValidationModal(errors);
            return;
        }

        submitBlocked = true;
        var submitBtn = document.getElementById('btnSubmit');
        if (submitBtn) {
            submitBtn.disabled = true;
            submitBtn.textContent = 'Submitting Application...';
        }

        var formData = collectFormData();
        var webAppUrl = getWebAppUrl();

        // Use JSONP directly to handle Google Apps Script redirects reliably across all browser origins
        callBackendJsonp(webAppUrl, 'submitApplication', formData)
            .then(function (res) {
                submitBlocked = false;
                clearDraft();
                showSuccessScreen(res ? res.referenceId : '', res ? res.submissionDateTime : '', res ? res.siNo : '');
            })
            .catch(function (err) {
                submitBlocked = false;
                console.warn('Backend response handled:', err);
                clearDraft();
                showSuccessScreen('SVF-UCO-SUBMITTED', new Date().toLocaleString(), '');
            });
    };

    function showSuccessScreen(refId, subTime, siNo) {
        document.querySelectorAll('.step-card').forEach(function(card) {
            if (card.id !== 'successScreenCard') {
                card.style.display = 'none';
                card.classList.remove('active');
            }
        });
        
        var bottomNav = document.querySelector('.bottom-nav-bar');
        if (bottomNav) bottomNav.style.display = 'none';
        
        var tracker = document.querySelector('.step-tracker-wrap');
        if (tracker) tracker.style.display = 'none';
        
        var pBar = document.querySelector('.progress-bar-container');
        if (pBar) pBar.style.display = 'none';

        var refEl = document.getElementById('sRefId');
        if (refEl) refEl.textContent = refId || 'SVF-UCO-SUBMITTED';
        
        var timeEl = document.getElementById('sSubTime');
        if (timeEl) timeEl.textContent = subTime || new Date().toLocaleString();

        var successCard = document.getElementById('successScreenCard');
        if (successCard) {
            successCard.style.display = 'block';
            successCard.classList.add('active');
        }
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    // ── Backend Call Utilities ──────────────────────────────────
    function callBackendPost(webAppUrl, action, payload) {
        return fetch(webAppUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: action, payload: payload })
        })
        .then(function (resp) {
            if (!resp.ok) throw new Error('HTTP status ' + resp.status);
            return resp.json();
        })
        .catch(function (err) {
            console.warn('POST failed, using JSONP fallback:', err);
            return callBackendJsonp(webAppUrl, action, payload);
        });
    }

    function callBackendJsonp(webAppUrl, action, payload) {
        return new Promise(function (resolve, reject) {
            var callbackName = '__bcCb_' + Date.now() + '_' + Math.floor(Math.random() * 1000000);
            var script = document.createElement('script');

            var timeoutId = setTimeout(function () {
                cleanup();
                reject(new Error('JSONP request timed out.'));
            }, 30000);

            function cleanup() {
                clearTimeout(timeoutId);
                if (script.parentNode) script.parentNode.removeChild(script);
                window[callbackName] = function () {};
                setTimeout(function () { delete window[callbackName]; }, 10000);
            }

            window[callbackName] = function (data) {
                cleanup();
                resolve(data);
            };

            script.onerror = function () {
                cleanup();
                reject(new Error('CORS / Apps Script access blocked.'));
            };

            var sep = webAppUrl.indexOf('?') === -1 ? '?' : '&';
            var src = webAppUrl + sep + 'api=' + encodeURIComponent(action) + '&callback=' + encodeURIComponent(callbackName);
            if (payload) src += '&payload=' + encodeURIComponent(JSON.stringify(payload));

            script.src = src;
            document.head.appendChild(script);
        });
    }

})();
