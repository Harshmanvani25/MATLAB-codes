%% Plasma Diagnostics Data Trimming (Robust Final Version - 2015a Compatible)

clc;
clear;
close all;

%% ================= USER CONFIGURATION =====================
FS = 1e5;                    % Sampling frequency (Hz)

IP_MIN_THRESHOLD   = 80;     % Minimum required peak IP
IGNORE_BEFORE_MS   = 60;     % Ignore plasma termination before this time
ZERO_WINDOW_MS     = 5;      % Required consecutive zero duration (ms)
EPSILON            = 0.5;    % Near-zero tolerance after ReLU
MIN_DURATION_MS    = 100;    % Minimum allowed plasma duration
MAX_DURATION_MS    = 130;    % Maximum allowed plasma duration

%% ================= SELECT OUTPUT BASE FOLDER ===============
OUTPUT_BASE_PATH = uigetdir(pwd, ...
    'Select folder to save all generated Excel files');

if OUTPUT_BASE_PATH == 0
    return;
end

fprintf('Trimmed files will be saved in:\n%s\n', OUTPUT_BASE_PATH);

%% ================= SELECT MULTIPLE EXCEL FILES ============
[fileNames, inputPath] = uigetfile( ...
    {'*.xlsx', 'Excel Files (*.xlsx)'}, ...
    'Select Excel File(s)', ...
    'MultiSelect', 'on');

if isequal(fileNames, 0)
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

    T = readtable(fullPath);

    varNamesOrig = T.Properties.VariableNames;
    varNamesLower = lower(varNamesOrig);

    %% ---- Detect Time column ------------------------------
    timeCandidates = {'time_ms','time','time(ms)','t'};
    timeIdx = false(size(varNamesLower));

    for k = 1:numel(timeCandidates)
        timeIdx = timeIdx | strcmp(varNamesLower, timeCandidates{k});
    end

    if ~any(timeIdx)
        numericCols = varfun(@isnumeric, T, 'OutputFormat','uniform');
        timeIdx = numericCols;
        timeIdx(2:end) = false;
    end

    if ~any(timeIdx)
        error('Time column not found.');
    end

    Time_ms = T{:, timeIdx};

    %% ---- Detect IP7 column -------------------------------
    ip7Idx = strcmp(varNamesOrig, 'IP7');

    if ~any(ip7Idx)
        error('IP7 column not found.');
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

    %% ---- Plasma START (First Time >= 0) ------------------
    start_idx = find(Time_ms >= 0, 1);

    if isempty(start_idx)
        fprintf('Skipped (No time >= 0)\n');
        continue;
    end

    start_time_ms = Time_ms(start_idx);

    %% ================= ROBUST END LOGIC ===================

    if max(IP7) < IP_MIN_THRESHOLD
        fprintf('Skipped (Peak IP < threshold)\n');
        continue;
    end

    IP_relu = max(IP7, 0);

    ZERO_WINDOW_SAMPLES = round((ZERO_WINDOW_MS/1000) * FS);

    ignore_idx = find(Time_ms >= IGNORE_BEFORE_MS, 1);
    if isempty(ignore_idx)
        fprintf('Skipped (No data beyond ignore window)\n');
        continue;
    end

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

    %% ---- Trim --------------------------------------------
    keepIdx = Time_ms >= start_time_ms & Time_ms <= end_time_ms;
    Time_ms_trim = Time_ms(keepIdx);

    N = numel(Time_ms_trim);
    Time_sec = (0:N-1).' / FS;
    Time_ms_fs = Time_sec * 1e3;

    plasma_duration_ms = N / FS * 1e3;

    if plasma_duration_ms < MIN_DURATION_MS || ...
       plasma_duration_ms > MAX_DURATION_MS
        fprintf('Skipped (Duration outside limits)\n');
        continue;
    end

    %% ---- Create output table -----------------------------
    T_out = table(Time_ms_trim, Time_ms_fs, ...
        'VariableNames', {'Time_ms_raw','Time_ms_fs'});

    for s = 1:numel(signalNames)
        T_out.(signalNames{s}) = T.(signalNames{s})(keepIdx);
    end

    %% ---- Rename Columns in Fixed Order -------------------
    requiredOrder = { ...
        'TIME','VLOOP2','IP7','OTCUR','HALPHA','CIII','OII', ...
        'HXR','SXR','HXRFLUX','BOLOV5','MIRNOV1','MIRNOV12'};

    if size(T_out,2) == numel(requiredOrder)
        T_out.Properties.VariableNames = requiredOrder;
    else
        warning('Column count mismatch. Renaming skipped.');
    end

    %% ---- Save --------------------------------------------
    [~, baseName, ~] = fileparts(fileName);
    outFile = [baseName '_new.xlsx'];
    outPath = fullfile(OUTPUT_BASE_PATH, outFile);

    writetable(T_out, outPath);

    fprintf('Saved: %s\n', outPath);
    fprintf('Plasma duration: %.2f ms\n', plasma_duration_ms);

    %% ================= PLOTTING ===========================

    numPlots = numel(signalNames);

    figure('Name', [baseName ' - Full Signals'], 'Color','w');

    for s = 1:numel(signalNames)

        subplot(numPlots,1,s)

        plot(Time_ms, T.(signalNames{s}));
        grid on
        ylabel(signalNames{s})

        hold on
        yl = ylim;
        plot([start_time_ms start_time_ms], yl, 'g--');
        plot([end_time_ms end_time_ms], yl, 'r--');
        hold off

    end

    figure('Name', [baseName ' - Plasma Only'], 'Color','w');

    for s = 1:numel(requiredOrder)
        subplot(numel(requiredOrder),1,s)
        plot(Time_ms_fs, T_out.(requiredOrder{s}));
        grid on
        ylabel(requiredOrder{s});
    end
    xlabel('Time (ms)');

end

disp('All files processed successfully.');
