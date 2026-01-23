%% Plasma Diagnostics Data Trimming (Auto Column Detection)
% IP7 auto-detected
% All signals trimmed identically
% Time-based logic enforced

clc;
clear;
close all;

%% ================= USER CONFIGURATION =====================
FS = 1e5;                    % Sampling frequency (Hz)
NEGATIVE_THRESHOLD = -3;     % IP7 threshold

%% ================= SELECT OUTPUT BASE FOLDER ===============
OUTPUT_BASE_PATH = uigetdir( ...
    pwd, ...
    'Select folder to save all generated Excel files');

if OUTPUT_BASE_PATH == 0
    disp('Output folder selection cancelled.');
    return;
end

fprintf('Output files will be saved in:\n%s\n', OUTPUT_BASE_PATH);


%% ================= SELECT MULTIPLE EXCEL FILES ============
[fileNames, inputPath] = uigetfile( ...
    {'*.xlsx', 'Excel Files (*.xlsx)'}, ...
    'Select Excel File(s)', ...
    'MultiSelect', 'on');

if isequal(fileNames, 0)
    disp('File selection cancelled');
    return;
end

if ischar(fileNames)
    fileNames = {fileNames};
end

%% ================= PROCESS FILES ==========================
for f = 1:numel(fileNames)


    fileName = fileNames{f};
    fullPath = fullfile(inputPath, fileName);
    fprintf('\nProcessing: %s\n', fileName);

    %% ---- Read table (single sheet) -----------------------
    T = readtable(fullPath);

    varNames = lower(T.Properties.VariableNames);

    %% ---- Detect Time column ------------------------------
    %% ---- Detect Time column (ROBUST) --------------------------
varNamesOrig = T.Properties.VariableNames;
varNames = lower(varNamesOrig);

% Preferred time names
timeCandidates = {'time_ms','time','time(ms)','t'};

timeIdx = false(size(varNames));
for k = 1:numel(timeCandidates)
    timeIdx = timeIdx | strcmp(varNames, timeCandidates{k});
end

% Fallback: first numeric column
if ~any(timeIdx)
    numericCols = varfun(@isnumeric, T, 'OutputFormat','uniform');
    timeIdx = numericCols;
    timeIdx(2:end) = false;   % first numeric column only
end

if ~any(timeIdx)
    error('Time column not found in %s', fileName);
end

Time_ms = T{:, timeIdx};

fprintf('Detected time column: %s\n', varNamesOrig{find(timeIdx,1)});

    %% ---- Detect IP7 column -------------------------------
    ip7Idx = contains(varNames, 'ip7');
    if ~any(ip7Idx)
        error('IP7 column not found in %s', fileName);
    end
    IP7 = T{:, ip7Idx};

    %% ---- Detect signal columns ---------------------------
    signalIdx = ~timeIdx & varfun(@isnumeric, T, 'OutputFormat','uniform');
    signalNames = T.Properties.VariableNames(signalIdx);

    %% ---- Sort by time (safety) ---------------------------
    [Time_ms, order] = sort(Time_ms);
    IP7 = IP7(order);

    for s = 1:numel(signalNames)
        T.(signalNames{s}) = T.(signalNames{s})(order);
    end

    %% ---- Plasma START (0 ms) -----------------------------
    start_idx = find(Time_ms == 0, 1);
    if isempty(start_idx)
        error('0 ms not found in %s', fileName);
    end
    start_time_ms = 0;

    %% ---- Plasma END (IP7 threshold) ----------------------
    end_rel = find(IP7(start_idx:end) <= NEGATIVE_THRESHOLD, 1);
    if isempty(end_rel)
        end_idx = numel(Time_ms);
    else
        end_idx = start_idx + end_rel - 1;
    end
    end_time_ms = Time_ms(end_idx);

    %% ---- Time-based trimming -----------------------------
    keepIdx = Time_ms >= start_time_ms & Time_ms <= end_time_ms;

    Time_ms_trim = Time_ms(keepIdx);

    %% ---- Rebuild DSP-accurate time -----------------------
    N = numel(Time_ms_trim);
    Time_sec = (0:N-1).' / FS;
    Time_ms_fs = Time_sec * 1e3;

    %% ---- Create output table dynamically -----------------
    T_out = table(Time_ms_trim, Time_ms_fs, ...
        'VariableNames', {'Time_ms_raw','Time_ms_fs'});

    for s = 1:numel(signalNames)
        sig = T.(signalNames{s})(keepIdx);
        T_out.(signalNames{s}) = sig;
    end

    %% ---- Save output -------------------------------------
    [~, baseName, ~] = fileparts(fileName);
    outFile = baseName + "_new.xlsx";
    outPath = fullfile(OUTPUT_BASE_PATH, outFile);
    writetable(T_out, outPath);

    %% ---- Plot (auto signals) -----------------------------
    figure('Name', baseName + " - Full Signals", 'Color','w');

    subplot(numel(signalNames),1,1)
    plot(Time_ms, IP7); hold on;
    xline(start_time_ms,'g--');
    xline(end_time_ms,'r--');
    yline(NEGATIVE_THRESHOLD,'k:');
    ylabel('IP7'); grid on;

    p = 2;
    for s = 1:numel(signalNames)
        if contains(lower(signalNames{s}), 'ip7'); continue; end
        subplot(numel(signalNames),1,p)
        plot(Time_ms, T.(signalNames{s}));
        xline(start_time_ms,'g--');
        xline(end_time_ms,'r--');
        ylabel(signalNames{s});
        grid on;
        p = p + 1;
    end
    sgtitle("Full Signals with Plasma Detection");

    %% ---- Plasma-only plots -------------------------------
    figure('Name', baseName + " - Plasma Only", 'Color','w');

    p = 1;
    for s = 1:numel(signalNames)
        subplot(numel(signalNames),1,p)
        plot(Time_ms_fs, T_out.(signalNames{s}));
        ylabel(signalNames{s});
        grid on;
        p = p + 1;
    end
    xlabel('Time (ms)');
    sgtitle("Plasma-Only Signals (Fs-Based Time)");

    %% ---- Summary -----------------------------------------
    fprintf('Saved: %s\n', outPath);
    fprintf('Plasma duration: %.3f ms\n', N/FS*1e3);
    fprintf('Samples kept   : %d\n', N);

end

disp('All files processed successfully.');

%% ================= DATASET GENERATION =====================

%% ================= SELECT FINAL TRIMMED FOLDER =============

finalFolder = uigetdir( ...
    "D:\IPR SEM 8 Internship\codes", ...
    'Select folder containing FINAL trimmed Excel files');

if finalFolder == 0
    disp('Folder selection cancelled');
    return;
end

excelFiles = dir(fullfile(finalFolder, '*_new.xlsx'));

if isempty(excelFiles)
    error('No *_new.xlsx files found in selected folder');
end

fprintf('Found %d trimmed Excel files\n', numel(excelFiles));
%% ================= DETECT SIGNAL COLUMNS ==================

sampleFile = fullfile(finalFolder, excelFiles(1).name);
T_sample = readtable(sampleFile);

varNamesOrig = T_sample.Properties.VariableNames;
varNamesLower = lower(varNamesOrig);

% Remove time columns
timeKeywords = {'time','time_ms','time_ms_raw','time_ms_fs'};
isTimeCol = false(1, numel(varNamesOrig));

for i = 1:numel(varNamesOrig)
    for k = 1:numel(timeKeywords)
        if contains(varNamesLower{i}, timeKeywords{k})
            isTimeCol(i) = true;
        end
    end
end

% Numeric columns only
isNumericCol = false(1, numel(varNamesOrig));
for i = 1:numel(varNamesOrig)
    isNumericCol(i) = isnumeric(T_sample.(varNamesOrig{i}));
end

signalIdx = isNumericCol & ~isTimeCol;
signalNames = varNamesOrig(signalIdx);

fprintf('\nDetected signal columns from FINAL trimmed data:\n');
disp(signalNames');
[selectedIdx, tf] = listdlg( ...
    'PromptString','Select signal(s) for dataset creation:', ...
    'ListString', signalNames, ...
    'SelectionMode','multiple');

if tf == 0
    disp('Signal selection cancelled');
    return;
end

selectedSignals = signalNames(selectedIdx);
%%
choice = questdlg( ...
    'Include Time column in dataset?', ...
    'Dataset Format', ...
    'With Time','Without Time','With Time');

includeTime = strcmp(choice,'With Time');

%% ================= SELECT DATASET OUTPUT FOLDER ============

datasetBaseFolder = uigetdir( ...
    "D:\IPR SEM 8 Internship\codes", ...
    'Select folder to save signal-wise datasets');

if datasetBaseFolder == 0
    disp('Dataset output folder selection cancelled');
    return;
end
%% ================= CREATE DATASETS =========================

for s = 1:numel(selectedSignals)

    sigName = selectedSignals{s};

    % Create signal-specific folder
    sigFolder = fullfile(datasetBaseFolder, sigName + "_all_excel_data");
    if ~exist(sigFolder,'dir')
        mkdir(sigFolder);
    end

    for f = 1:numel(excelFiles)

        fileName = excelFiles(f).name;
        fullPath = fullfile(finalFolder, fileName);

        T = readtable(fullPath);

        % Safety check
        if ~ismember(sigName, T.Properties.VariableNames)
            warning('%s missing in %s', sigName, fileName);
            continue;
        end

        % Prepare output table
        if includeTime
            Tout = table( ...
                T.Time_ms_fs, ...
                T.(sigName), ...
                'VariableNames', {'Time_ms', sigName});
        else
            Tout = table( ...
                T.(sigName), ...
                'VariableNames', {sigName});
        end

        % Output filename
        baseName = erase(fileName, "_new.xlsx");
        outFile = baseName + "_" + sigName + ".xlsx";

        % Save
        writetable(Tout, fullfile(sigFolder, outFile));
    end
end

disp('Signal-wise dataset creation completed successfully.');

