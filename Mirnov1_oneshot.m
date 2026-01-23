%% Interactive Excel Import Script

clc;
clear;

%% Step 1: Select Excel file
[fileName, folderPath] = uigetfile( ...
    {'*.xlsx;*.xls', 'Excel Files (*.xlsx, *.xls)'}, ...
    'Select Excel File');

% If user cancels
if isequal(fileName, 0)
    disp('File selection cancelled.');
    return;
end

fullFilePath = fullfile(folderPath, fileName);

%% Step 2: Get sheet names from the selected Excel file
[~, sheetNames] = xlsfinfo(fullFilePath);

if isempty(sheetNames)
    error('No sheets found in the selected Excel file.');
end

%% Step 3: Let user select sheet
[sheetIndex, tf] = listdlg( ...
    'PromptString', 'Select Sheet:', ...
    'SelectionMode', 'single', ...
    'ListString', sheetNames);

if tf == 0
    disp('Sheet selection cancelled.');
    return;
end

selectedSheet = sheetNames{sheetIndex};

%% Step 4: Read data from selected sheet
opts = detectImportOptions(fullFilePath, 'Sheet', selectedSheet);
dataTable = readtable(fullFilePath, opts);

%% Step 5: Create variable name from Excel filename
[~, baseFileName, ~] = fileparts(fileName);
validVarName = matlab.lang.makeValidName(baseFileName);

%% Step 6: Assign data to variable with Excel filename
assignin('base', validVarName, table2array(dataTable));

%% Done
fprintf('Data imported successfully!\n');
fprintf('Variable name: %s\n', validVarName);
fprintf('Sheet used: %s\n', selectedSheet);
