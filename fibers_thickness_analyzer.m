% Fiber Thickness Analyzer
% ------------------------------------------------------------
% Author: Tatiana Pisarenko
% Email: Tatiana.Pisarenko@vut.cz
% Created: 14-Aug-2025
% MATLAB Version: R2025a
%
% Description:
%   Interactive tool for detecting, measuring, and analyzing
%   fiber thickness in microscopy images. Supports manual
%   scale calibration, live segmentation tuning, histogram
%   visualization, and export of results to Excel.
%
% Usage:
%   Run the script from the Editor tab or enter its name in
%   the Command Window. MATLAB's current folder must contain
%   this script file.
%
% Changelog:
%   2025-08-14  v1.0  First functional release.
% ------------------------------------------------------------

function fibers_thickness_analyzer
    clear; clc; close all;

    %% --- 1) File selection ---
    [filename, pathname] = uigetfile( ...
        {'*.tif;*.tiff;*.png;*.jpg;*.jpeg', 'Image files'}, ...
        'Select an image');
    if isequal(filename, 0), error('No file selected.'); end
    filePath = fullfile(pathname, filename);

    info = imfinfo(filePath);
    if numel(info) > 1
        sliceNum = input(sprintf('The file has %d frames. Select (1–%d): ', ...
            numel(info), numel(info)));
        I = imread(filePath, sliceNum);
    else
        I = imread(filePath);
    end

    %% --- 2) Grayscale image ---
    Igray = toGray2D(I);
    Igray = im2uint8(Igray);

    %% --- 3) Manual scale input ---
    prompt = {'Physical scale length:', 'Unit (nm/µm/mm):', 'Scale in pixels:'};
    def    = {'10', 'µm', '100'};
    answ   = inputdlg(prompt, 'Scale settings', [1 40], def);
    if isempty(answ), error('Cancelled by user.'); end

    physLen_val = str2double(answ{1});
    unit        = lower(strtrim(answ{2}));
    pxLen       = str2double(answ{3});

    if isnan(physLen_val) || isnan(pxLen) || pxLen <= 0
        error('Invalid scale values.');
    end

    switch unit
        case {'um', 'µm', 'μm'}
            physLen_um = physLen_val;
        case {'nm'}
            physLen_um = physLen_val / 1000;
        case {'mm'}
            physLen_um = physLen_val * 1000;
        otherwise
            error('Unsupported unit: %s', unit);
    end
    scale_um_per_px = physLen_um / pxLen;

    %% --- 4) Interactive GUI ---
    f = figure('Name', 'Fiber Thickness Analyzer', 'NumberTitle', 'off', ...
               'Position', [100 100 1200 600]);

    panelHeight = 0.12;
    ctrlPanel = uipanel('Parent', f, 'Units', 'normalized', ...
        'Position', [0 0 1 panelHeight], 'Title', 'Controls');

    % Layout parameters
    imgBottom  = panelHeight + 0.12;
    imgHeight  = 0.8 - imgBottom;
    histBottom = imgBottom;
    histHeight = imgHeight;

    axImg = subplot('Position', [0.05 imgBottom 0.4 imgHeight]);
    hImg  = imshow(Igray, 'Parent', axImg);
    title(axImg, 'Detected Fibers');

    axHist = subplot('Position', [0.55 histBottom 0.4 histHeight]);

    % Default parameters
    params.sens        = 0.5;
    params.cannyThresh = 0.2;
    params.minThick    = 0;      % min fiber thickness (µm)
    params.maxThick    = inf;    % max fiber thickness (µm)
    binScale           = 1.0;

    %% --- Controls ---
    % Sensitivity
    uicontrol('Parent', ctrlPanel, 'Style', 'text', 'Units', 'normalized', ...
        'Position', [0.05 0.55 0.1 0.3], ...
        'String', sprintf('Sensitivity: %.2f', params.sens), ...
        'Tag', 'txtSens');
    s1 = uicontrol('Parent', ctrlPanel, 'Style', 'slider', ...
        'Min', 0.01, 'Max', 1, 'Value', params.sens, ...
        'Units', 'normalized', 'Position', [0.05 0.15 0.1 0.3], ...
        'Callback', @(~, ~) updateSeg());

    % Canny Threshold
    uicontrol('Parent', ctrlPanel, 'Style', 'text', 'Units', 'normalized', ...
        'Position', [0.20 0.55 0.1 0.3], ...
        'String', sprintf('Canny Threshold: %.2f', params.cannyThresh), ...
        'Tag', 'txtCanny');
    s2 = uicontrol('Parent', ctrlPanel, 'Style', 'slider', ...
        'Min', 0.01, 'Max', 0.8, 'Value', params.cannyThresh, ...
        'Units', 'normalized', 'Position', [0.20 0.15 0.1 0.3], ...
        'Callback', @(~, ~) updateSeg());

    % Min fiber thickness
    uicontrol('Parent', ctrlPanel, 'Style', 'text', 'Units', 'normalized', ...
        'Position', [0.35 0.55 0.1 0.3], ...
        'String', sprintf('Min. thickness: %.2f µm', params.minThick), ...
        'Tag', 'txtMinThick');
    s3 = uicontrol('Parent', ctrlPanel, 'Style', 'slider', ...
        'Min', 0, 'Max', 1, 'Value', params.minThick, ...
        'Units', 'normalized', 'Position', [0.35 0.15 0.1 0.3], ...
        'SliderStep', [0.001 0.001], ...
        'Callback', @(~, ~) updateHist());

    % Max fiber thickness
    uicontrol('Parent', ctrlPanel, 'Style', 'text', 'Units', 'normalized', ...
        'Position', [0.50 0.55 0.1 0.3], ...
        'String', sprintf('Max. thickness: %.2f µm', params.maxThick), ...
        'Tag', 'txtMaxThick');
    s4 = uicontrol('Parent', ctrlPanel, 'Style', 'slider', ...
        'Min', 0.01, 'Max', 10, 'Value', 10, ...
        'Units', 'normalized', 'Position', [0.50 0.15 0.1 0.3], ...
        'SliderStep', [0.001 0.001], ...
        'Callback', @(~, ~) updateHist());

    % Bin scale
    uicontrol('Parent', ctrlPanel, 'Style', 'text', 'Units', 'normalized', ...
        'Position', [0.65 0.55 0.1 0.3], ...
        'String', sprintf('Bin scale: %.2fx', binScale), ...
        'Tag', 'txtBinScale');
    s5 = uicontrol('Parent', ctrlPanel, 'Style', 'slider', ...
        'Min', 0.5, 'Max', 5, 'Value', binScale, ...
        'Units', 'normalized', 'Position', [0.65 0.15 0.1 0.3], ...
        'SliderStep', [0.05 0.05], ...
        'Callback', @(~, ~) updateHist());

    % Export button
    uicontrol('Parent', ctrlPanel, 'Style', 'pushbutton', ...
        'String', 'Export to Excel', 'Units', 'normalized', ...
        'Position', [0.85 0.3 0.1 0.4], ...
        'Callback', @(~, ~) exportData());

    fiberThickness_um = [];

    updateSeg(); % initial display

    %% --- Segmentation & measurement ---
    function updateSeg()
        params.sens        = get(s1, 'Value');
        params.cannyThresh = get(s2, 'Value');

        set(findobj(f, 'Tag', 'txtSens'), ...
            'String', sprintf('Sensitivity: %.2f', params.sens));
        set(findobj(f, 'Tag', 'txtCanny'), ...
            'String', sprintf('Canny Threshold: %.2f', params.cannyThresh));

        Ieq  = adapthisteq(Igray, 'ClipLimit', 0.01);
        Iflt = imgaussfilt(Ieq, 1.2);

        lowT  = params.cannyThresh;
        highT = min(lowT + 0.2, 0.99);
        if lowT >= highT, lowT = highT - 0.01; end
        lowT = max(0.01, lowT);

        BW_adapt = imbinarize(Iflt, 'adaptive', 'Sensitivity', params.sens);
        E = edge(Iflt, 'Canny', [lowT highT]);
        E = imdilate(E, strel('disk', 1));
        E_filled = imfill(E, 'holes');

        BW = BW_adapt | E_filled;
        BW = imopen(BW, strel('disk', 1));
        BW = imclose(BW, strel('disk', 2));
        BW = bwareaopen(BW, 30);

        skel     = bwmorph(BW, 'skel', Inf);
        D        = bwdist(~BW);
        diam_px  = 2 * D;

        CC = bwconncomp(BW);
        fiberThickness_um = nan(CC.NumObjects, 1);
        for i = 1:CC.NumObjects
            mask = false(size(BW));
            mask(CC.PixelIdxList{i}) = true;
            skelMask = mask & skel;
            if any(skelMask(:))
                diamVals_um = diam_px(skelMask) * scale_um_per_px;
                fiberThickness_um(i) = mean(diamVals_um);
            end
        end

        outline = imdilate(bwperim(BW), strel('disk', 1));
        overlay = imoverlay(Igray, outline, [1 1 0]);
        set(hImg, 'CData', overlay);
        title(axImg, sprintf('Fibers detected: %d', numel(fiberThickness_um)));

        updateHist();
    end

    %% --- Histogram update ---
    function updateHist()
        params.minThick = get(s3, 'Value');
        params.maxThick = get(s4, 'Value');
        binScale        = get(s5, 'Value');

        set(findobj(f, 'Tag', 'txtMinThick'), ...
            'String', sprintf('Min. thickness: %.2f µm', params.minThick));
        set(findobj(f, 'Tag', 'txtMaxThick'), ...
            'String', sprintf('Max. thickness: %.2f µm', params.maxThick));
        set(findobj(f, 'Tag', 'txtBinScale'), ...
            'String', sprintf('Bin scale: %.2fx', binScale));

        cla(axHist);
        if ~isempty(fiberThickness_um)
            filtered = fiberThickness_um( ...
                fiberThickness_um >= params.minThick & ...
                fiberThickness_um <= params.maxThick);
            if ~isempty(filtered)
                rangeVals = max(filtered) - min(filtered);
                autoBins  = max(5, min(30, round(rangeVals / 0.5)));
                numBins   = max(1, round(autoBins * binScale));
                histogram(axHist, filtered, numBins);
            end
        end
        xlabel(axHist, 'Fiber thickness [µm]');
        ylabel(axHist, 'Count');
        title(axHist, 'Thickness Distribution');
        grid(axHist, 'on');
        axHist.XTickLabelRotation = 30;
    end

    %% --- Data export ---
    function exportData()
        if isempty(fiberThickness_um)
            warndlg('No data to export.', 'Export');
            return;
        end
        [file, path] = uiputfile('fiber_data.xlsx', 'Save Excel file');
        if isequal(file, 0), return; end
        T = table((1:numel(fiberThickness_um))', fiberThickness_um, ...
            'VariableNames', {'ID', 'thickness_um'});
        writetable(T, fullfile(path, file));
        msgbox('Data successfully exported.', 'Export');
    end
end

%% --- Grayscale conversion helper ---
function Igray = toGray2D(I)
    if ndims(I) == 2
        Igray = I;
    elseif ndims(I) == 3
        if size(I, 3) == 3
            Igray = rgb2gray(I);
        else
            Igray = I(:, :, 1);
        end
    else
        Igray = I(:, :, 1);
    end
end
