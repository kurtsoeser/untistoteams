
const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

ns.updateTeacherStats = function updateTeacherStats() {
    const uniqueTeachers = new Set(ns.filteredData.map(row => row.lehrer.toUpperCase().trim()));
    const teachersArray = Array.from(uniqueTeachers);
    const mappedCount = teachersArray.filter(t => ns.teacherEmailMapping[t]).length;
    const unmappedCount = teachersArray.length - mappedCount;

    document.getElementById('uniqueTeachersNeeded').textContent = teachersArray.length;
    document.getElementById('mappedTeachers').textContent = mappedCount;
    document.getElementById('unmappedTeachers').textContent = unmappedCount;
    document.getElementById('teacherRequiredStats').style.display = 'grid';

    if (unmappedCount > 0) ns.displayMissingTeachers(teachersArray);
    else document.getElementById('missingTeachersSection').style.display = 'none';

    if (Object.keys(ns.teacherEmailMapping).length > 0) {
        ns.displayTeacherMappingTableWithUsage(teachersArray);
    }
};

ns.displayMissingTeachers = function displayMissingTeachers(allTeachers) {
    const unmappedTeachers = allTeachers.filter(t => !ns.teacherEmailMapping[t]);
    if (unmappedTeachers.length === 0) {
        document.getElementById('missingTeachersSection').style.display = 'none';
        return;
    }
    const emailDomain =
        typeof window.ms365GetTeacherEmailDomainSuffix === 'function'
            ? window.ms365GetTeacherEmailDomainSuffix()
            : '@';
    const tbody = document.getElementById('missingTeachersBody');
    tbody.replaceChildren();
    unmappedTeachers.forEach(kuerzel => {
        const generatedEmail = kuerzel.toLowerCase() + emailDomain;
        const tr = document.createElement('tr');
        const td1 = document.createElement('td');
        const strong = document.createElement('strong');
        strong.textContent = kuerzel;
        td1.appendChild(strong);
        const td2 = document.createElement('td');
        td2.textContent = generatedEmail;
        const td3 = document.createElement('td');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-small';
        btn.textContent = '➕ Hinzufügen';
        btn.addEventListener('click', () => ns.quickAddTeacher(kuerzel, generatedEmail));
        td3.appendChild(btn);
        tr.append(td1, td2, td3);
        tbody.appendChild(tr);
    });
    document.getElementById('missingTeachersSection').style.display = 'block';
};

ns.quickAddTeacher = function quickAddTeacher(kuerzel, suggestedEmail) {
    ns.openModal(
        'E-Mail für ' + kuerzel,
        '<label for="quickEmail">E-Mail-Adresse</label><input type="email" id="quickEmail" value="' +
            ns.attrEscape(suggestedEmail) +
            '">',
        () => {
            const email = document.getElementById('quickEmail').value.trim().toLowerCase();
            if (!email) {
                ns.showToast('Bitte eine E-Mail eingeben.');
                return;
            }
            ns.teacherEmailMapping[kuerzel] = email;
            document.getElementById('teacherCount').textContent = Object.keys(ns.teacherEmailMapping).length;
            ns.closeModal();
            ns.updateTeacherStats();
            document.getElementById('teacherMappingInfo').style.display = 'block';
        }
    );
};

ns.goToStep = function goToStep(rawStep) {
    const panel = document.getElementById('panelWebuntis');
    if (!panel) return;

    let step = parseInt(String(rawStep).trim(), 10);
    if (!Number.isFinite(step) || step < 0 || step > 8) return;

    // Nur ab Schritt 6 (Graph/CSV/Schüler): generierte Teams nötig.
    // Schritt 5 („Teams konfigurieren“) ist bewusst ausgenommen — dort wird erst generiert.
    if (step >= 6 && step <= 8) {
        const validTeams = ns.teamsData.filter(t => t.isValid);
        if (!ns.teamsGenerated || validTeams.length === 0) {
            ns.showToast('Bitte zuerst unter „Teams konfigurieren“ auf „Team-Namen generieren“ klicken (mindestens ein gültiges Team).');
            step = 5;
        }
    }

    panel.querySelectorAll('.step-content').forEach(el => el.classList.remove('active'));
    panel.querySelectorAll('#panelWebuntis .steps > .step').forEach(el => {
        el.classList.remove('active');
        el.classList.remove('completed');
    });

    const contentEl = panel.querySelector('.step-content[data-step="' + step + '"]');
    const tabEl = panel.querySelector('#panelWebuntis .steps > .step[data-step="' + step + '"]');
    if (!contentEl || !tabEl) return;

    contentEl.classList.add('active');
    tabEl.classList.add('active');
    try {
        tabEl.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
    } catch (e) {
        /* ignore */
    }

    const stepOrder = [0, 1, 2, 3, 4, 5, 6, 7, 8];
    const currentIndex = stepOrder.indexOf(step);
    if (currentIndex >= 0) {
        for (let i = 0; i < currentIndex; i++) {
            const prev = panel.querySelector('#panelWebuntis .steps > .step[data-step="' + stepOrder[i] + '"]');
            if (prev) prev.classList.add('completed');
        }
    }

    ns.currentStep = step;

    if (step === 1) {
        const seeded =
            typeof ns.seedWebuntisPasteIfEmpty === 'function' ? ns.seedWebuntisPasteIfEmpty() : false;
        if (seeded) {
            ns.showToast('Demo: 6 Beispielzeilen aus Schul‑Standards vorbelegt.');
        }
    }

    const hint = document.getElementById('manualKursteamHint');
    if (hint) hint.style.display = step === 2 && ns.kursteamEntryMode === 'manual' ? 'block' : 'none';

    if (step === 3) {
        if (typeof ns.displayEditableData === 'function') ns.displayEditableData();
        if (typeof ns.displayManualTeamsPreview === 'function') ns.displayManualTeamsPreview();
    }
    if (step === 4) {
        ns.updateTeacherStats();
        const btnNextTeamCfg = document.getElementById('continueBtn3');
        if (btnNextTeamCfg) btnNextTeamCfg.style.display = '';
    }
    if (step === 5) {
        const manRow = document.getElementById('kursteamManualAddRow');
        if (manRow) manRow.style.display = ns.teamsGenerated ? '' : 'none';
    }
    if (step === 8) {
        if (typeof ns.seedStudentRosterFromTenantIfEmpty === 'function') {
            const seeded = ns.seedStudentRosterFromTenantIfEmpty();
            if (seeded === 'demo') ns.showToast('Demo: Schülerliste aus Schul‑Standards vorbelegt.');
            else if (seeded === 'tenant') ns.showToast('Schülerliste aus Schul‑Einstellungen übernommen.');
        }
        if (typeof ns.refreshStudentRosterUI === 'function') ns.refreshStudentRosterUI();
    }
    if (step === 7) ns.prepareCSVExport();

    const stepsBar = panel.querySelector('.steps');
    if (typeof window.ms365ApplyStepProgress === 'function') {
        window.ms365ApplyStepProgress(stepsBar, step, stepOrder);
    }
};

ns.prepareCSVExport = function prepareCSVExport() {
    const validTeams = ns.teamsData.filter(t => t.isValid);
    document.getElementById('exportCount').textContent = validTeams.length;

    const warn = document.getElementById('step4NoTeamsWarning');
    const ready = document.getElementById('step4ReadyHint');
    const dl = document.getElementById('btnDownloadCsv');
    if (validTeams.length === 0) {
        warn.style.display = 'block';
        ready.style.display = 'none';
        dl.disabled = true;
    } else {
        warn.style.display = 'none';
        ready.style.display = 'block';
        dl.disabled = false;
    }

    let csvPreview = ns.buildCsvRow(['TeamName', 'Gruppenmail', 'Besitzer']);
    validTeams.slice(0, 5).forEach(team => {
        csvPreview += ns.buildCsvRow([team.teamName, team.gruppenmail, team.besitzer]);
    });
    if (validTeams.length > 5) {
        csvPreview += '... (' + (validTeams.length - 5) + ' weitere Teams)\n';
    }
    document.getElementById('csvPreview').textContent = csvPreview;
};

ns.downloadCSV = function downloadCSV() {
    const validTeams = ns.teamsData.filter(t => t.isValid);
    if (validTeams.length === 0) {
        ns.showToast('Keine gültigen Teams zum Exportieren.');
        return;
    }
    let csv = ns.buildCsvRow(['TeamName', 'Gruppenmail', 'Besitzer']);
    validTeams.forEach(team => {
        csv += ns.buildCsvRow([team.teamName, team.gruppenmail, team.besitzer]);
    });
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'neueteams.csv';
    link.click();
    URL.revokeObjectURL(link.href);
};

ns.copyPowerShell = function copyPowerShell() {
    const script = document.getElementById('powershellScript').textContent;
    navigator.clipboard.writeText(script).then(() => {
        ns.showToast('PowerShell-Script in die Zwischenablage kopiert.');
    });
};

function buildStandaloneKursteamPs1(validTeams) {
    const stamp = new Date().toISOString();
    const rows = validTeams.map(t =>
        "    [PSCustomObject]@{ TeamName = '" +
            ns.psEscapeSingle(t.teamName) +
            "'; Gruppenmail = '" +
            ns.psEscapeSingle(t.gruppenmail) +
            "'; Besitzer = '" +
            ns.psEscapeSingle(t.besitzer) +
            "' }"
    );
    const loginBlock = [
        'Write-Host ""',
        'Write-Host "=== Anmeldung bei Microsoft Teams / Microsoft 365 ===" -ForegroundColor Cyan',
        'Write-Host "Konten mit MFA: bitte Option A waehlen (Browser-Anmeldung)." -ForegroundColor Yellow',
        'Write-Host ""',
        'Write-Host " [A] Interaktive Anmeldung (empfohlen, MFA moeglich)"',
        'Write-Host " [B] Benutzername + Passwort (Get-Credential) – oft nur ohne MFA zuverlaessig"',
        'Write-Host ""',
        '$loginChoice = Read-Host "Auswahl eingeben (A oder B, Standard A)"',
        'if ($loginChoice -eq "B" -or $loginChoice -eq "b") {',
        '    $script:Ms365Cred = Get-Credential -Message "Microsoft 365 / Teams Administrator"',
        '    if ($null -eq $script:Ms365Cred) { Write-Error "Anmeldung abgebrochen."; exit 1 }',
        '    Connect-MicrosoftTeams -Credential $script:Ms365Cred',
        '} else {',
        '    Connect-MicrosoftTeams',
        '}',
        ''
    ].join('\r\n');

    const lines = [];
    lines.push('#Requires -Version 5.1');
    lines.push('# Kursteam-Anlage (Microsoft Teams, Vorlage EDU_Class)');
    lines.push('# Entspricht weiterhin Microsoft Learn: New-Team -Template "EDU_Class" (gueltige Werte: EDU_Class, EDU_PLC).');
    lines.push('# Microsoft empfiehlt fuer Klassen-Teams das Modul MicrosoftTeams in Version 7.3.1 oder neuer.');
    lines.push('# Erzeugt in der Browser-App am ' + stamp);
    lines.push('# Daten sind unten eingebettet – keine separate CSV noetig.');
    lines.push('');
    lines.push('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8');
    lines.push('$ErrorActionPreference = "Continue"');
    lines.push('');
    lines.push('if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {');
    lines.push('    Write-Host "Installiere Modul MicrosoftTeams (einmalig)..." -ForegroundColor Yellow');
    lines.push('    Install-Module MicrosoftTeams -Scope CurrentUser -Force');
    lines.push('}');
    lines.push('Import-Module MicrosoftTeams -ErrorAction Stop');
    lines.push('');
    lines.push(loginBlock);
    lines.push('$TeamsList = @(');
    lines.push(rows.join(',\r\n'));
    lines.push(')');
    lines.push('');
    lines.push('$i = 0');
    lines.push('foreach ($Team in $TeamsList) {');
    lines.push('    $i++');
    lines.push('    try {');
    lines.push(
        '        $null = New-Team -Template "EDU_Class" -DisplayName $Team.TeamName -MailNickName $Team.Gruppenmail -Owner $Team.Besitzer -ErrorAction Stop'
    );
    lines.push('        Write-Host ("OK [{0}/{1}] {2}" -f $i, $TeamsList.Count, $Team.Gruppenmail) -ForegroundColor Green');
    lines.push('    }');
    lines.push('    catch {');
    lines.push('        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $Team.Gruppenmail, $_.Exception.Message)');
    lines.push('    }');
    lines.push('    Start-Sleep -Seconds 2');
    lines.push('}');
    lines.push('');
    lines.push('Write-Host ""');
    lines.push('Write-Host "Fertig. Fenster schliesst nicht automatisch." -ForegroundColor Cyan');
    lines.push('Read-Host "Enter druecken zum Beenden"');
    return lines.join('\r\n');
}

ns.downloadKursteamStandalonePackage = function downloadKursteamStandalonePackage() {
    const validTeams = ns.teamsData.filter(t => t.isValid);
    if (!validTeams.length) {
        ns.showToast('Keine gültigen Teams – zuerst Team-Namen generieren.');
        return;
    }
    if (typeof window.ms365BuildPolyglotCmd !== 'function') {
        ns.showToast('polyglot-cmd.js fehlt – Seite neu laden.');
        return;
    }
    const ps1 = buildStandaloneKursteamPs1(validTeams);
    const cmd = window.ms365BuildPolyglotCmd({
        title: 'Kursteam-Anlage',
        echoLine: 'Starte Kursteam-Anlage mit PowerShell ...',
        psBody: ps1
    });
    ns.downloadBlob('Kursteam-Anlage.cmd', cmd);
    ns.showToast('Kursteam-Anlage.cmd heruntergeladen – Doppelklick zum Start.');
};

ns.resetApp = function resetApp() {
    ns.confirmModal('App zurücksetzen', 'Alle Daten in dieser Sitzung wirklich verwerfen? (Lokaler Zwischenstand bleibt, bis Sie ihn löschen.)', () => {
        location.reload();
    });
};

// Step-Header klickbar + keyboard
document.querySelectorAll('#panelWebuntis .steps > .step').forEach(step => {
    step.setAttribute('tabindex', '0');
    step.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            step.click();
        }
    });
    step.addEventListener('click', function () {
        const stepNum = parseInt(String(this.dataset.step).trim(), 10);
        const currentStepNum = parseInt(String(ns.currentStep).trim(), 10);
        const current = Number.isFinite(currentStepNum) ? currentStepNum : 0;
        if (!Number.isFinite(stepNum) || stepNum < 0 || stepNum > 8) return;
        if (stepNum <= current || this.classList.contains('completed')) {
            ns.goToStep(stepNum);
        }
    });
});

// Global exports für HTML onclick
window.goToStep = ns.goToStep;
window.downloadCSV = ns.downloadCSV;
window.copyPowerShell = ns.copyPowerShell;
window.resetApp = ns.resetApp;
window.downloadKursteamStandalonePackage = ns.downloadKursteamStandalonePackage;

// Snapshot für Microsoft Graph im Browser (kursteam-graph.js).
window.ms365GetKursteamSnapshotForGraph = function () {
    const validTeams = ns.teamsData.filter(t => t.isValid);
    if (!validTeams.length) return null;
    return {
        teams: validTeams.map(t => ({
            teamName: t.teamName,
            gruppenmail: t.gruppenmail,
            besitzer: String(t.besitzer || '').trim()
        }))
    };
};

document.addEventListener('DOMContentLoaded', () => {
    const panel = document.getElementById('panelWebuntis');
    if (!panel || typeof window.ms365ApplyStepProgress !== 'function') return;
    const order = [0, 1, 2, 3, 4, 5, 6, 7, 8];
    const parsed = parseInt(String(ns.currentStep).trim(), 10);
    const step = Number.isFinite(parsed) ? parsed : 0;
    window.ms365ApplyStepProgress(panel.querySelector('.steps'), step, order);
});

