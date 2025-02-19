% script octave/matlab for timbre analysis
% written3/2023 by Malte Kob
clear all
pkg load signal % only needed in octave to load the signal package
pkg load statistics % only needed in octave to load the signal package
pkg load io % only needed in octave to load and save xlsx files

function [subDirsNames] = GetSubDirsFirstLevelOnly(parentDir)
    % Get a list of all files and folders in this folder.
    files = dir(parentDir);
    % Get a logical vector that tells which is a directory.
    dirFlags = [files.isdir];
    % Extract only those that are directories.
    subDirs = files(dirFlags); % A structure with extra info.
    % Get only the folder names into a cell array.
    subDirsNames = {subDirs(3:end).name};
end

FontSize=9;
fmin     = 0;    % minimum frequency for display range
fmax     = 4000;% maximum frequency for display rang
fRange1L = fmin; % lower limit of 1st frequency range
fRange1U = 500; % upper limit of 1st frequency range
fRange2L = 500; % lower limit of 2nd frequency range
fRange2U = 1500; % upper limit of 2nd frequency range
fRange3L = 1500; % lower limit of 3rd frequency range
fRange3U = fmax; % upper limit of 3rd frequency range
fcutoff   = 100; % Cutoff frequency for highpass filter to filter out noise
%timemin = 0.2; % start of time signal%
%timemax = 1.5; % end of time signal
timeoffset   = 0.1; % start of evaluation of signal
timeduration = 2.0; % desired length of evaluated time signal

SPLmax    = 80;
sens      =  43.7; % sensitivity in mv/Pa
dispRange = 110;
yref      = 1;

%tpath = 'C:\Users\krist\OneDrive\Documents\CS_Pt_2\Programs\'; % path with praat.exe and temporary wav file
tpath = 'c:\ProgrammeOhneInstallation\';
tfname = 'tempraatfile.wav';
%fpath  = 'C:\Users\krist\OneDrive\Desktop\Analysis_Selections\';
fpath  = 'c:\Users\Malte Kob\Filr\Für mich freigegeben\Audio_Analysis_Selections\';
%fname  = '1.a.1.wav';

cd(fpath);
%partlist = dir;
partlist = GetSubDirsFirstLevelOnly(pwd)
for partidx = 1:size(partlist,2)
  actpart = char(partlist(partidx))
  cd(actpart)
  figpath='C:\temp\';
%  IMAGE_PATH = figpath;
  reclist = dir('*.wav');
  for recidx = 1:size(reclist,1)
    fname_org = reclist(recidx).name;
    fname=[fname_org(1:end-4)];
    fname=strrep(fname,'.','_')
    fnamestr=strrep(fname,'_','\_');
    fnamecellarray=cellstr(fname);

    [amplituderaw,FS] = audioread(fname_org);
    amplituderaw = amplituderaw(:,1); % take first channel of multichannel wav files

    % Octave
    n = 1;
    sf2=FS/2;
    Wc = fcutoff/sf2;
    [b, a] = butter (n, Wc, "high"); % filter with butterworth highpass filter
    amplitudeHP = filter(b,a,amplituderaw);

    amplitudenf = amplituderaw;
    amplitudewf = amplitudeHP;
    width = 1/8; %fast = 1/8 second, slow = 1 second
    rc = 5e-3; % rise constant (?)

    % create RMS values
    [wx w]   = movingrms (zscore (amplitudenf),width,rc,FS);
    [wxf wf] = movingrms (zscore (amplitudewf),width,rc,FS);

    % create SPL values
    SPL=20*log10(wx) + SPLmax;
    SPLf=20*log10(wxf) + SPLmax;

    % detect onset and offset

    % cut signal to constant length
    idxmax     = length(amplitudenf)-1;
    timevecraw = [0:idxmax]/FS;

    midwx = (max(wxf)-min(wxf))/2;
    onsetidx = min(find(wxf>midwx));
    offsetidx = max(find(wxf>midwx));
    timeoffsetidx = round(FS*timeoffset);
    timeminidx = onsetidx + timeoffsetidx;
    timemaxidx = timeminidx+round(FS*timeduration);
    if timemaxidx > offsetidx-timeoffsetidx
      timemaxidx = offsetidx-timeoffsetidx;
      maxTime = (timemaxidx-timeminidx)/FS;
      disp(["The signal is too short for an analysis of ", num2str(timeduration), " s. Stattdessen wurden ", num2str(maxTime)," s ausgewertet."])
    else
      maxTime = timeoffset+timeduration;
    end

    %timeminidx = min(find(timevecraw>timemin));
    %timemaxidx = max(find(timevecraw<timemax));
    timevec    = timevecraw(timeminidx:timemaxidx)';
    amplitude  = amplitudenf(timeminidx:timemaxidx);
    amplitudef  = amplitudewf(timeminidx:timemaxidx);
    ampwin=hanning(length(amplitude));
    %amplitude  = amplitude .* ampwin;
    audiowrite ([tpath,tfname], amplitude, FS)
    %sound(amplitude,FS,24);
    player=audioplayer(amplitude,FS);play(player);
    %dosstr = [tpath,'praat.exe --run ',tpath,'TesStcipt.praat'];
    dosstr = [tpath,'praat.exe --run ',tpath,'FormantAnalysisScript.praat'];
    [STATUS,TEXTraw]=dos(dosstr);
    if STATUS == 0
      TEXT = '';
      for chari =1:length(TEXTraw)
        icharacter = str2double(TEXTraw(chari));
        if isnan(icharacter)
          if TEXTraw(chari) == '.'
            TEXT = [TEXT,'.'];
          elseif TEXTraw(chari) == ';'
            TEXT = [TEXT,';'];
          endif
        else
          TEXT = [TEXT,num2str(icharacter)];
        endif
      endfor
      FORMANTS = str2num(TEXT);
      printf ("Praat call successful! - ");
    else
      printf ("Praat call not successful! - ");
    endif


    % temporal display
    timefigh=figure(1);

    subplot(2,1,1)
    plot(timevecraw,amplitudenf/sens*1000, timevecraw,amplitudewf/sens*1000, timevec,amplitude/sens*1000, timevec,amplitudef/sens*1000)
    xlabel('Time (s)')
    ylabel('Sound pressure (Pa)')
    %axis([timemin timemax -inf inf])
    grid on
    title([actpart,' - ',fnamestr,'.wav']);
    legend('before selection and filtering')
    set(gca,"Fontsize",FontSize)

    subplot(2,1,2)
    plot(timevecraw,SPL,'r',timevecraw,SPLf,'b')
    %set(gcf,'Position',[50 250 1200 800])
    hold on
    area (timevecraw(timeminidx:timemaxidx), SPL(timeminidx:timemaxidx),'facecolor',[0.85 0.85 1],'edgecolor','m','linewidth',2);
    area (timevecraw(timeminidx:timemaxidx), SPLf(timeminidx:timemaxidx),'facecolor',[0.45 0.45 0.85],'edgecolor','c','linewidth',2);
    hold off
    xlabel('Time (s)')
    ylabel('SPL (dB)')
    %axis([timemin timemax SPLmax-80 SPLmax])
    grid on
    legend('unfiltered', 'filtered')
    set(gca,"Fontsize",FontSize)

    % Spectrogram
    ##subplot(3,1,3) % spectrogram of signal
    ##sliceWidth = 20; % slice width in ms
    ##windowWidth = 100; % window width in ms


    % Matlab
    %Nx = timemaxidx;
    %nsc = floor(Nx/4.5);
    %ns = 8; % sections of spectrogram
    %ov = 0.5; % overlap of sections
    %lsc = floor(Nx/(ns-(ns-1)*ov));
    %nov = floor(nsc/2);
    %nff = max(256,2^nextpow2(nsc));
    %t = spectrogram(amplitude,lsc,floor(ov*lsc),nff);
    % spectrogram(amplitudef,128*16,120,128*32,FS,'yaxis'); % Matlab
    %axis([0 timemax 0 fmax/1000])


    % Octave
    ##step=ceil(sliceWidth*FS/1000);    % one spectral slice every 20 ms
    ##window=ceil(windowWidth*FS/1000); % 100 ms data window
    ##specgram(amplitudef, 2^nextpow2(window), FS, window, window-step);
    ##axis([0 timemax-timemin 0 fmax])
    ##%xlabel('Frequency (kHz)')
    ##%ylabel('SPL (dB)')
    set(gca,"Fontsize",FontSize)

    %savefig(timefigh,[figpath,'time_',fname])
    saveas(timefigh,[figpath,'time_',fname,'.png'])


    freqfigh=figure(2); % Spektrum of selected signal
    clf
    % Octave
    N = timemaxidx-timeminidx+1; % number of data points
    spec = fft(amplitude); % numerical approx. of FT
    specf = fft(amplitudef); % numerical approx. of FT
    df = FS/N; % spacing between samples on freq. axis
    min_f = -FS/2; % min freq. for which fft is calculated
    max_f = FS/2 - df; % max freq. for which fft is calculated
    f = [min_f : df : max_f]'; % horizontal values
    size(f); % should equal N
    y = 20*log10(abs(fftshift(spec))); % level of shifted spectrum
    yf = 20*log10(abs(fftshift(specf))); % level of shifted spectrum
    Lmin = max(y)-dispRange;
    Lmax = max(y)+5;

    plot(f, y, 'r', f, yf, 'b')
    %set(gcf,'Position',[100 200 1200 800])
    xlabel('Frequency (Hz)')
    ylabel('FFT (dB)')
    axis([0 fmax Lmin Lmax])

    hold on
    line([FORMANTS(1),FORMANTS(1)],[Lmin+5 Lmax-3], 'marker', 'd');
    text(FORMANTS(1),Lmin+3,'F1', 'fontsize', FontSize)
    fu=FORMANTS(1)-FORMANTS(6)/2;fo=FORMANTS(1)+FORMANTS(6)/2;
    fuidx=min(find(f>fu));foidx=max(find(f<fo));
    yfu=yf(fuidx);yfo=yf(foidx);
    area (f(fuidx:foidx), y(fuidx:foidx),'facecolor',[0.85 0.85 1],'edgecolor',[0.85 0.85 1],'linewidth',2);
    line([FORMANTS(2),FORMANTS(2)],[Lmin+5 Lmax-3], 'marker', 'd');
    text(FORMANTS(2),Lmin+3,'F2', 'fontsize', FontSize)
    fu=FORMANTS(2)-FORMANTS(7)/2;fo=FORMANTS(2)+FORMANTS(7)/2;
    fuidx=min(find(f>fu));foidx=max(find(f<fo));
    yfu=yf(fuidx);yfo=yf(foidx);
    area (f(fuidx:foidx), y(fuidx:foidx),'facecolor',[0.85 0.85 1],'edgecolor',[0.85 0.85 1],'linewidth',2);
    line([FORMANTS(3),FORMANTS(3)],[Lmin+5 Lmax-3], 'marker', 'd');
    text(FORMANTS(3),Lmin+3,'F3', 'fontsize', FontSize)
    fu=FORMANTS(3)-FORMANTS(8)/2;fo=FORMANTS(3)+FORMANTS(8)/2;
    fuidx=min(find(f>fu));foidx=max(find(f<fo));
    yfu=yf(fuidx);yfo=yf(foidx);
    area (f(fuidx:foidx), y(fuidx:foidx),'facecolor',[0.85 0.85 1],'edgecolor',[0.85 0.85 1],'linewidth',2);
    line([FORMANTS(4),FORMANTS(4)],[Lmin+5 Lmax-3], 'marker', 'd');
    text(FORMANTS(4),Lmin+3,'F4', 'fontsize', FontSize)
    fu=FORMANTS(4)-FORMANTS(9)/2;fo=FORMANTS(4)+FORMANTS(9)/2;
    fuidx=min(find(f>fu));foidx=max(find(f<fo));
    yfu=yf(fuidx);yfo=yf(foidx);
    area (f(fuidx:foidx), y(fuidx:foidx),'facecolor',[0.85 0.85 1],'edgecolor',[0.85 0.85 1],'linewidth',2);
    line([FORMANTS(5),FORMANTS(5)],[Lmin+5 Lmax-3], 'marker', 'd');
    text(FORMANTS(5),Lmin+3,'F5', 'fontsize', FontSize)
    fu=FORMANTS(5)-FORMANTS(10)/2;fo=FORMANTS(5)+FORMANTS(10)/2;
    fuidx=min(find(f>fu));foidx=max(find(f<fo));
    yfu=yf(fuidx);yfo=yf(foidx);
    area (f(fuidx:foidx), y(fuidx:foidx),'facecolor',[0.85 0.85 1],'edgecolor',[0.85 0.85 1],'linewidth',2);

    hold off

    grid on
    legend('unfiltered', 'filtered', 'Formant frequency', 'formant bandwidth')
    %legend('Formants from Praat','formant bandwidth', 'unfiltered', 'filtered')
    set(gca,"Fontsize",FontSize)
    title([actpart,' - ',fnamestr,'.wav']);

    %savefig(freqfigh,[figpath,'freq_',fname])
    saveas(freqfigh,[figpath,'freq_',fname,'.png'])
    FormantTable = [[1:5]',FORMANTS(1:5),FORMANTS(6:10)]; % first formant frequencies, then formant bandwidths
    rstatus = xlswrite ([figpath,'Formants_',fname], FormantTable);



    ERh=figure(3); % Calculation of STER

    clf

    % For each phrase, the mean acoustic power in dB (unweighted) was computed.
    % We then computed the average power spectral density by means of the FFT
    % calculated on a series of overlapping segments throughout the duration of
    % the phrase. The segment size was set to 20 ms (440 samples), with segments
    % spaced at 5 ms. A Blackman window was applied to each segment which was then
    % extended to 2048 samples before computing its FFT. The mean of the squared
    % absolute FFTs was calculated and normalized to take account of the effects
    % of the window. The power in the average spectrum is thus comparable to the
    % power obtained from the time domain acoustic signal. From the average power
    % spectral density for each phrase, the power in the frequency bands 0�2 kHz
    % and 2�4 kHz was obtained and denoted by Plo and Phi , respectively.
    % The choice of these frequency bands was based on previous studies that have
    % shown that the acoustic energy or peak amplitude within the band 2�4 kHz
    % gives a good representation of the �ringing� quality in a singer�s voice.
    % The ratio Phi /Plo was also calculated.

    % reference settings
    subplot(3,1,1) % spectrogram of signal with reference settings
    sliceWidth = 5; % slice distance in ms std: 20ms vs. 5ms
    windowWidth = 20; % window width in ms; std: 100ms vx. 20ms
    step=ceil(sliceWidth*FS/1000);    % one spectral slice every 20 ms
    window=ceil(windowWidth*FS/1000); % 100 ms data window
    %SpecH=specgram(amplitudef, 2^nextpow2(window), FS, window, window-step);
    fftn = 2^nextpow2(window); % next highest power of 2
    [SpecH, fspec, tspec] = specgram(amplitudef, fftn, FS, window, window-step);
    %S = abs(SpecH(2:fftn*fmax/FS,:)); % magnitude in range 0<f<=fmax Hz.
    S = abs(SpecH); % magnitude in range 0<f<=fmax Hz.
    S = S/max(S(:)); % normalize magnitude so that max is 0 dB.
    S = max(S, 10^(-40/10)); % clip below -40 dB.
    %S = min(S, 10^(-3/10)); % clip above -3 dB.
    imagesc (tspec, fspec, log(S)); % display in log scale
    set (gca, "ydir", "normal"); % put the 'y' direction in the correct direction
    axis([timeoffset maxTime fRange1L fRange3U])
    title([actpart,' - ',fnamestr,'.wav']);
    xlabel('Time (s)')
    ylabel('Frequency (Hz)')
    set(gca,"Fontsize",FontSize)

    subplot(3,1,2) % split Spectrum
    maxTidx = size(SpecH,2);
    %timeVecRaw = [0:maxTidx-1]/maxTidx*timemax;
    maxFidx = size(SpecH,1);
    freqVecRaw = [0:maxFidx-1]/maxFidx*FS/2;
    fRange1Lidx = min(find(freqVecRaw>=fRange1L)); % lower limit of 1st frequency range
    fRange1Uidx = max(find(freqVecRaw<fRange1U)); % upper limit of 1st frequency range
    fRange2Lidx = min(find(freqVecRaw>=fRange2L)); % lower limit of 2nd frequency range
    fRange2Uidx = max(find(freqVecRaw<fRange2U)); % upper limit of 2nd frequency range
    fRange3Lidx = min(find(freqVecRaw>=fRange3L)); % lower limit of 3rd frequency range
    fRange3Uidx = max(find(freqVecRaw<=fRange3U)); % upper limit of 3rd frequency range

    idx = round(maxTidx/2); % middle slice in time for spectrum
    idx = 50; % arbitrary index of slice in time
    ER1 = mean(20*log10(abs(SpecH(fRange1Lidx:fRange1Uidx,:))),2);
    ER2 = mean(20*log10(abs(SpecH(fRange2Lidx:fRange2Uidx,:))),2);
    ER3 = mean(20*log10(abs(SpecH(fRange3Lidx:fRange3Uidx,:))),2);
    plot(freqVecRaw(fRange1Lidx:fRange1Uidx),ER1,'g',freqVecRaw(fRange2Lidx:fRange2Uidx),ER2,'m',freqVecRaw(fRange3Lidx:fRange3Uidx),ER3,'c')
    %set(gcf,'Position',[150 150 1200 800])
    axis([fRange1L fRange3U max(ER1)-dispRange max(ER1)+5])
    grid on
    title('Filtered Spectrum')
    xlabel('Frequency (Hz)')
    ylabel('SPL (dB)')
    legend(['Range 1: ',num2str(fRange1L),'-',num2str(fRange1U),' Hz'; 'Range 2: ',num2str(fRange2L),'-',num2str(fRange2U),' Hz'; 'Range 3: ',num2str(fRange3L),'-',num2str(fRange3U),' Hz'])
    set(gca,"Fontsize",FontSize)


    subplot(3,1,3) % STEF of the signal
    STER1 = mean(20*log10(abs(SpecH(fRange1Lidx:fRange1Uidx,:))),1)';
    STER2 = mean(20*log10(abs(SpecH(fRange2Lidx:fRange2Uidx,:))),1)';
    STER3 = mean(20*log10(abs(SpecH(fRange3Lidx:fRange3Uidx,:))),1)';
    STER12 = STER1-STER2;
    STER32 = STER3-STER2;
    STER31 = STER3-STER1;
    %plot(timeVecRaw,STER1,'g',timeVecRaw,STER2,'m')
    plot(tspec,STER1,'g',tspec,STER2,'m',tspec,STER3,'c')
    grid on
    ##title('Short-Term Energy')
    xlabel('Time (s)')
    ylabel('Short-Term Energy (dB)')
    set(gca,"Fontsize",FontSize)
    ax = gca;
    yl = get(ax,'ylim');
    yt = get(ax,'ytick');
    h0 = get(ax,'children');
    hold on
    [ax,h1,h2] = plotyy(ax,0,0,tspec,STER12); % plot for Chiaoscuro Kristen
    [ax,h1,h2] = plotyy(ax,0,0,tspec,STER32); % plot for Chiaoscuro Kristen
    [ax,h1,h2] = plotyy(ax,0,0,tspec,STER31); % plot for Chiaoscuro Kristen
    hold off
    delete(h1)
    set(h2,'color','k')
    ylabel (ax(2), "'Short-Term Energy ratio (dB)'")
    legend('STE Range 1', 'STE Range 2', 'STE Range 3', 'STER 1-2 (R1-R2) (dB)', 'STER 3-2 (R3-R2) (dB)', 'STER 3-1 (R3-R1) (dB)')
    set(gca,"Fontsize",FontSize)
    hold off

    %savefig(ERh,[figpath,'ER_',fname])
    saveas(ERh,[figpath,'ER_',fname,'.png'])


    boxh=figure(4);

    %boxplot({STER21,STER32,STER31}, {'STER21', 'STER32', 'STER31'});
    boxstatistics=boxplot({STER12,STER32,STER31},"Labels", {'STER 1-2', 'STER 3-2', 'STER 3-1'});
    %set(gcf,'Position',[200 100 1200 800])

    meanSTER1   = mean(STER1);
    medianSTER1 = median(STER1);
    stdSTER1    = std(STER1);
    meanSTER2   = mean(STER2);
    medianSTER2 = median(STER2);
    stdSTER2    = std(STER2);
    meanSTER3   = mean(STER3);
    medianSTER3 = median(STER3);
    stdSTER3    = std(STER3);
    %header1=['Mean STER1',num2str(meanSTER1);'Mean STER2',num2str(meanSTER2);;'Mean STER3',num2str(meanSTER3);];
    %header2=['Median STER1',num2str(medianSTER1);'Median STER2',num2str(medianSTER2);;'Median STER3',num2str(medianSTER3);];
    %header2=['Std STER1',num2str(stdSTER1);'Std STER2',num2str(stdSTER2);;'Std STER3',num2str(stdSTER3);];

    ERTable=[meanSTER1,medianSTER1,stdSTER1;meanSTER2,medianSTER2,stdSTER2;...
    meanSTER3,medianSTER3,stdSTER3];

    ExcelContent = [[1:13];FormantTable,[ERTable;0,0,0;0,0,0],[boxstatistics,zeros(7,2)]'];

    title([actpart,' - ',fnamestr,'.wav']);
    ylabel('STER')
    set(gca,'ylim',[-40 40])
    %set(gca,'xlim',[0 4])
    grid on
    set(gca,"Fontsize",FontSize)
    %savefig(boxh,[figpath,'ERBox_',fname])
    saveas(boxh,[figpath,'ERBox_',fname,'.png'])
    rstatus = xlswrite ([figpath,'ERBoxStats_',fname], boxstatistics);
    rstatus = xlswrite ([figpath,'AllStats_',fname], ExcelContent);

    ExcelRowStart = (recidx-1)*6 + 1;
    ExcelRowStartStr = num2str(ExcelRowStart);
    ExcelRowEnd   = ExcelRowStart + 5;
    ExcelRowEndStr= num2str(ExcelRowEnd);
    ExcelStrRec = ['A',ExcelRowStartStr,':A',ExcelRowStartStr];
    ExcelStrTab = ['B',ExcelRowStartStr,':N',ExcelRowEndStr]
    rstatus = xlswrite ([figpath,'AllStats'], fnamecellarray, actpart, ExcelStrRec);
    rstatus = xlswrite ([figpath,'AllStats'], ExcelContent, actpart, ExcelStrTab);

    % AllStats Excel Table
    % 1st row: Index of Parameter
    % 1st col: Index of Formant (1..5), STE range (1..3), STER (1-2, 3-2, 3-1)
    % 2nd col: Formant Frequency (Hz)
    % 3rd col: Formant Bandwidth (Hz)
    % 4th col: ST Energy mean (dB)
    % 5th col: ST Energy median (dB)
    % 6th col: ST Energy standard deviation (dB)
    % 7th col: STER minimum (dB)
    % 8th col: STER 1st quartile (dB)
    % 9th col: STER 2nd quartile (median) (dB)
    % 10th col: STER 3rd quartile (dB)
    % 11th col: STER Maximum (dB)
    % 12th col: STER Lower confidence limit for median (dB)
    % 13th col: STER Upper confidence limit for median (dB)
  endfor
  cd ..
endfor

