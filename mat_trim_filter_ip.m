%% Plasma Diagnostics Data Trimming (Robust Logic Version)
% IP7 auto-detected
% Sustained zero-window collapse detection
% Configurable parameters
% Trimmed files saved in selected output folder

clc;
clear;
close all;

%% ================= USER CONFIGURATION =====================
FS = 1e5;                    % Sampling frequency (Hz)

% -------- Plasma validity filter ----------
IP_MIN_THRESHOLD   = 80;     % Minimum required peak IP
IGNORE_BEFORE_MS   = 60;     % Ignore plasma termination before this time

% -------- Zero-window detection ----------
ZERO_WINDOW_MS     = 5;      % Required consecutive zero duration (ms)
EPSILON            = 0.5;    % Near-zero tolerance after ReLU

% -------- Duration validation -------------
MIN_DURATION_MS    = 100;    % Accept plasma only if duration >= this
MAX_DURATION_MS    = 130;    % Accept plasma only if duration <= this

%% ================= SELECT OUTPUT BASE FOLDER ===============
OUTPUT_BASE_PATH = uigetdir(pwd, ...
    'Select folder to save all generated Excel files');

if OUTPUT_BASE_PATH == 0
    disp('Output folder selection cancelled.');
    return;
end

fprintf('Trimmed files will be saved in:\n%s\n', OUTPUT_BASE_PATH);

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

    %% ---- Read table --------------------------------------
    T = readtable(fullPath);

    varNamesOrig = T.Properties.VariableNames;
    varNames = lower(varNamesOrig);

    %% ---- Detect Time column ------------------------------
    timeCandidates = {'time_ms','time','time(ms)','t'};
    timeIdx = false(size(varNames));

    for k = 1:numel(timeCandidates)
        timeIdx = timeIdx | strcmp(varNames, timeCandidates{k});
    end

    if ~any(timeIdx)
        numericCols = varfun(@isnumeric, T, 'OutputFormat','uniform');
        timeIdx = numericCols;
        timeIdx(2:end) = false;
    end

    if ~any(timeIdx)
        error('Time column not found in %s', fileName);
    end

    Time_ms = T{:, timeIdx};
    fprintf('Detected time column: %s\n', ...
        varNamesOrig{find(timeIdx,1)});

    %% ---- Detect IP7 column -------------------------------
    ip7Idx = contains(varNames, 'ip7');
    if ~any(ip7Idx)
        error('IP7 column not found in %s', fileName);
    end

    IP7 = T{:, ip7Idx};

    %% ---- Detect signal columns ---------------------------
    signalIdx = ~timeIdx & varfun(@isnumeric, T, 'OutputFormat','uniform');
    signalNames = T.Properties.VariableNames(signalIdx);

    %% ---- Sort by time ------------------------------------
    [Time_ms, order] = sort(Time_ms);
    IP7 = IP7(order);

    for s = 1:numel(signalNames)
        T.(signalNames{s}) = T.(signalNames{s})(order);
    end

    %% ---- Plasma START (0 ms) -----------------------------
    start_idx = find(Time_ms == 0, 1);
    if isempty(start_idx)
        fprintf('Skipped (0 ms not found)\n');
        continue;
    end

    start_time_ms = 0;

    %% ======================================================
    %% ROBUST PLASMA END LOGIC
    %% ======================================================

    % 1) Validate minimum peak current
    if max(IP7) < IP_MIN_THRESHOLD
        fprintf('Skipped (Peak IP < %.1f)\n', IP_MIN_THRESHOLD);
        continue;
    end

    % 2) ReLU transform
    IP_relu = max(IP7, 0);

    % 3) Convert window to samples
    ZERO_WINDOW_SAMPLES = round((ZERO_WINDOW_MS/1000) * FS);

    % 4) Ignore early region
    ignore_idx = find(Time_ms >= IGNORE_BEFORE_MS, 1);
    if isempty(ignore_idx)
        fprintf('Skipped (No data beyond %.1f ms)\n', IGNORE_BEFORE_MS);
        continue;
    end

    % 5) Detect sustained near-zero window
    end_idx = numel(Time_ms);
    zero_counter = 0;

    for i = ignore_idx:numel(IP_relu)

        if IP_relu(i) <= EPSILON
            zero_counter = zero_counter + 1;
        else
            zero_counter = 0;
        end

        if zero_counter >= ZERO_WINDOW_SAMPLES
            end_idx = i - ZERO_WINDOW_SAMPLES + 1;
            break;
        end
    end

    end_time_ms = Time_ms(end_idx);

    %% ---- Time-based trimming -----------------------------
    keepIdx = Time_ms >= start_time_ms & Time_ms <= end_time_ms;
    Time_ms_trim = Time_ms(keepIdx);

    %% ---- Rebuild DSP-accurate time -----------------------
    N = numel(Time_ms_trim);
    Time_sec = (0:N-1).' / FS;
    Time_ms_fs = Time_sec * 1e3;

    plasma_duration_ms = N / FS * 1e3;

    %% ---- Duration validation -----------------------------
    if plasma_duration_ms < MIN_DURATION_MS || ...
       plasma_duration_ms > MAX_DURATION_MS

        fprintf('Skipped (Duration %.2f ms outside [%d,%d])\n', ...
            plasma_duration_ms, MIN_DURATION_MS, MAX_DURATION_MS);
        continue;
    end

    %% ---- Create output table -----------------------------
    T_out = table(Time_ms_trim, Time_ms_fs, ...
        'VariableNames', {'Time_ms_raw','Time_ms_fs'});

    for s = 1:numel(signalNames)
        T_out.(signalNames{s}) = T.(signalNames{s})(keepIdx);
    end

    %% ---- Save trimmed file -------------------------------
    [~, baseName, ~] = fileparts(fileName);
    outFile = baseName + "_new.xlsx";
    outPath = fullfile(OUTPUT_BASE_PATH, outFile);

    writetable(T_out, outPath);

    fprintf('Saved: %s\n', outPath);
    fprintf('Plasma duration: %.2f ms\n', plasma_duration_ms);
    fprintf('Samples kept   : %d\n', N);

end

disp('All files processed successfully.');

%% ==========================================================
%% DATASET GENERATION (UNCHANGED FROM YOUR ORIGINAL)
%% ==========================================================

finalFolder = uigetdir(pwd, ...
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

sampleFile = fullfile(finalFolder, excelFiles(1).name);
T_sample = readtable(sampleFile);

varNamesOrig = T_sample.Properties.VariableNames;
varNamesLower = lower(varNamesOrig);

timeKeywords = {'time','time_ms','time_ms_raw','time_ms_fs'};
isTimeCol = false(1, numel(varNamesOrig));

for i = 1:numel(varNamesOrig)
    for k = 1:numel(timeKeywords)
        if contains(varNamesLower{i}, timeKeywords{k})
            isTimeCol(i) = true;
        end
    end
end

isNumericCol = false(1, numel(varNamesOrig));
for i = 1:numel(varNamesOrig)
    isNumericCol(i) = isnumeric(T_sample.(varNamesOrig{i}));
end

signalIdx = isNumericCol & ~isTimeCol;
signalNames = varNamesOrig(signalIdx);

[selectedIdx, tf] = listdlg( ...
    'PromptString','Select signal(s) for dataset creation:', ...
    'ListString', signalNames, ...
    'SelectionMode','multiple');

if tf == 0
    disp('Signal selection cancelled');
    return;
end

selectedSignals = signalNames(selectedIdx);

choice = questdlg( ...
    'Include Time column in dataset?', ...
    'Dataset Format', ...
    'With Time','Without Time','With Time');

includeTime = strcmp(choice,'With Time');

datasetBaseFolder = uigetdir(pwd, ...
    'Select folder to save signal-wise datasets');

if datasetBaseFolder == 0
    disp('Dataset output folder selection cancelled');
    return;
end

for s = 1:numel(selectedSignals)

    sigName = selectedSignals{s};
    sigFolder = fullfile(datasetBaseFolder, sigName + "_all_excel_data");

    if ~exist(sigFolder,'dir')
        mkdir(sigFolder);
    end

    for f = 1:numel(excelFiles)

        fileName = excelFiles(f).name;
        fullPath = fullfile(finalFolder, fileName);

        T = readtable(fullPath);

        if ~ismember(sigName, T.Properties.VariableNames)
            continue;
        end

        if includeTime
            Tout = table(T.Time_ms_fs, T.(sigName), ...
                'VariableNames', {'Time_ms', sigName});
        else
            Tout = table(T.(sigName), ...
                'VariableNames', {sigName});
        end

        baseName = erase(fileName, "_new.xlsx");
        outFile = baseName + "_" + sigName + ".xlsx";

        writetable(Tout, fullfile(sigFolder, outFile));
    end
end

disp('Signal-wise dataset creation completed successfully.');
