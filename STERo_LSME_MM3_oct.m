% script octave/matlab for timbre analysis
% written3/2023 by Malte Kob
clear all
pkg load signal % only needed in octave to load the signal package
pkg load statistics % only needed in octave to load the signal package
FontSize=10;
%figpath = 'd:\LSME\Auswertung\STER\';
%path = 'c:\Users\Malte Kob\Filr\Meine Dateien\Projekte\LSME\NTi\';
%path = 'c:\Users\Malte Kob\Filr\Meine Dateien\Projekte\Lingwave-Validierung\NTi\MM\'; % Michaela
figpath = 'c:\temp\1a';


%Auswertung Prob. 724/2008/vowels/mittlere 1 sec
%path = 'D:\LSME\Rohdaten\LingWAVES_backup\clients\lwintern9\';
%vowel = 'a' ; name = 'lwintern9_20081022_111648.wav'; timemin=3,753; timemax=4,753;
%vowel = 'e' ; name = 'lwintern9_20081022_111648.wav'; timemin= 6,313; timemax= 7,313;
%vowel = 'i' ; name = 'lwintern9_20081022_111648.wav'; timemin= 8,914; timemax= 9,914;
%vowel = 'o' ; name = 'lwintern9_20081022_111648.wav'; timemin= 11,41; timemax= 12,41;
%vowel = 'u' ; name = 'lwintern9_20081022_111648.wav'; timemin= 14,027; timemax= 15,0265;


%Auswertung Prob. 724/2012/vowels/mittlere 1 sec
%path = 'D:\LSME\Rohdaten\LingWAVES_backup\clients\lwintern9\';
%vowel = 'a' ; name = ' lwintern9_20120618_110031.wav'; timemin= 0,596; timemax= 1,596;
%vowel = 'e' ; name = ' lwintern9_20120618_110031.wav'; timemin= 2,4815; timemax= 3,4815;
%vowel = 'i' ; name = ' lwintern9_20120618_110031.wav'; timemin= 4,8185; timemax= 5,8185;
%vowel = 'o' ; name = ' lwintern9_20120618_110031.wav'; timemin= 7,445; timemax= 8,445;
%vowel = 'u' ; name = ' lwintern9_20120618_110031.wav'; timemin= 10,077; timemax= 11,0765;

%Auswertung Prob. 724/2023/vowels/mittlere 1 sec
path = 'c:\temp\1a\Audio\';
%path = 'D:\LSME\Rohdaten\LingWAVES_backup\clients\lwintern9\';
vowel = 'a' ;
name = '3.d.4-Take1.wav';
timemin = 4.3525;
timemax = 5.3525;
%vowel = 'a' ; name = 'lwintern9_20230320_153337.wav'; timemin= 4,3525; timemax= 5,3525;
%vowel = 'e' ; name = 'lwintern9_20230320_153337.wav'; timemin= 12,024; timemax= 13,024;
%vowel = 'i' ; name = 'lwintern9_20230320_153337.wav'; timemin= 19,682; timemax= 20,682;
%vowel = 'o' ; name = 'lwintern9_20230320_153337.wav'; timemin= 27,918; timemax= 28,918;
%vowel = 'u' ; name = 'lwintern9_20230320_153337.wav'; timemin= 35,281; timemax= 36,281;



%timemin = 1.7; % start of time signal

% Ende gegeben:
%timemax = 4.2; % end of time signal
duration = timemax-timemin; % length of extraced signal

% Dauer gegeben:
%duration = 3; % length of extraced signal
%timemax = timemin + duration; % extract of 2 seconds


fmax = 10000;    % maximum frequency for display
fRange1L = 0;   % lower limit of 1st frequency range
fRange1U = 2000; % upper limit of 1st frequency range
fRange2L = 2000; % lower limit of 2nd frequency range
fRange2U = 4000; % upper limit of 2nd frequency range
SPLmax=129.2;
sens = 43.7; % sensitivity in mv/Pa
dispRange = 110;
yref= 1;
filename = [path,name];
[amplituderaw,FS] = audioread(filename);
idxmax = length(amplituderaw)-1;
timevecraw = [0:idxmax]/FS;
fcutoff = 100; % Cutoff frequency for highpass filter
tmin = 0;

timeminidx = min(find(timevecraw>timemin));
timemaxidx = max(find(timevecraw<timemax));
timevec = timevecraw(timeminidx:timemaxidx)';
amplitude = amplituderaw(timeminidx:timemaxidx);
ampwin=hanning(length(amplitude));
amplitude = amplitude .* ampwin;
sound(amplitude,FS,24);

% filter signal to remove fan noise
% Matlab
%amplitudef = highpass(amplitude, fcutoff, FS); % remove low frequency noise

% Octave
n = 1;
sf2=FS/2;
Wc = fcutoff/sf2;
[b, a] = butter (n, Wc, "high"); % filter with butterworth highpass filter
amplitudeHP = filter(b,a,amplitude);

%cWeight = weightingFilter('C-weighting','SampleRate',FS);
%amplitudec = cWeight(amplitude);

%aWeight = weightingFilter('A-weighting','SampleRate',FS);
%amplitudea = aWeight(amplitude);

amplitude = amplitude;
amplitudef = amplitudeHP;

SPL=10*log10(amplitude.*amplitude) + SPLmax;
SPLf=10*log10(amplitudef.*amplitudef) + SPLmax;


% temporal display
figure(1)
subplot(3,1,1)
plot(timevecraw,amplituderaw/sens*1000)
%detectSpeech(amplitudef,FS, "Thresholds",thresholdAverage);
xlabel('Time (s)')
ylabel('Sound pressure (Pa)')
%axis([timemin timemax -inf inf])
grid on
legend('before selection and filtering')
set(gca,"Fontsize",FontSize)

subplot(3,1,2)
plot(timevec,SPL,'r',timevec,SPLf,'b')
xlabel('Time (s)')
ylabel('SPL (dB)')
axis([timemin timemax SPLmax-80 SPLmax])
grid on
legend('unfiltered', 'filtered')
set(gca,"Fontsize",FontSize)

% Spectrogram
subplot(3,1,3) % spectrogram of signal
sliceWidth = 20; % slice width in ms
windowWidth = 100; % window width in ms


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
step=ceil(sliceWidth*FS/1000);    % one spectral slice every 20 ms
window=ceil(windowWidth*FS/1000); % 100 ms data window
specgram(amplitudef, 2^nextpow2(window), FS, window, window-step);
axis([0 timemax-timemin 0 fmax])
%xlabel('Frequency (kHz)')
%ylabel('SPL (dB)')
set(gca,"Fontsize",FontSize)


figure(2) % Spektrum of selected signal
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
plot(f, y, 'r', f, yf, 'b')
xlabel('Frequency (Hz)')
ylabel('FFT (dB)')
axis([0 fmax max(y)-dispRange max(y)+5])
grid on
legend('unfiltered', 'filtered')
set(gca,"Fontsize",FontSize)

figure(3) % Calculation of STER
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
[SpecH, f, t] = specgram(amplitudef, fftn, FS, window, window-step);
%S = abs(SpecH(2:fftn*fmax/FS,:)); % magnitude in range 0<f<=fmax Hz.
S = abs(SpecH); % magnitude in range 0<f<=fmax Hz.
S = S/max(S(:)); % normalize magnitude so that max is 0 dB.
S = max(S, 10^(-40/10)); % clip below -40 dB.
%S = min(S, 10^(-3/10)); % clip above -3 dB.
imagesc (t, f, log(S)); % display in log scale
set (gca, "ydir", "normal"); % put the 'y' direction in the correct direction
axis([0 timemax-timemin 0 fRange2U])
title('Filtered signal')
xlabel('Time (s)')
ylabel('Frequency (Hz)')
set(gca,"Fontsize",FontSize)

subplot(3,1,2) % split Spectrum
maxTidx = size(SpecH,2);
timeVecRaw = [0:maxTidx-1]/maxTidx*timemax;
maxFidx = size(SpecH,1);
freqVecRaw = [0:maxFidx-1]/maxFidx*FS/2;
fRange1Lidx = min(find(freqVecRaw>fRange1L)); % lower limit of 1st frequency range
fRange1Uidx = max(find(freqVecRaw<fRange1U)); % upper limit of 1st frequency range
fRange2Lidx = min(find(freqVecRaw>fRange2L)); % lower limit of 2nd frequency range
fRange2Uidx = max(find(freqVecRaw<fRange2U)); % upper limit of 2nd frequency range

idx = round(maxTidx/2); % middle slice in time for spectrum
idx = 50; % arbitrary index of slice in time
ER1 = mean(20*log10(abs(SpecH(fRange1Lidx:fRange1Uidx,:))),2);
ER2 = mean(20*log10(abs(SpecH(fRange2Lidx:fRange2Uidx,:))),2);
plot(freqVecRaw(fRange1Lidx:fRange1Uidx),ER1,'g',freqVecRaw(fRange2Lidx:fRange2Uidx),ER2,'m')
axis([fRange1L fRange2U max(ER1)-dispRange max(ER1)+5])
grid on
title('Filtered Spectrum')
xlabel('Frequency (Hz)')
ylabel('Sound pressure Level (dB)')
legend(['Range 1: ',num2str(fRange1L),'-',num2str(fRange1U),' Hz'], ['Range 2: ',num2str(fRange2L),'-',num2str(fRange2U),' Hz'])
set(gca,"Fontsize",FontSize)


subplot(3,1,3) % STEF of the signal
STER1 = mean(20*log10(abs(SpecH(fRange1Lidx:fRange1Uidx,:))),1)';
STER2 = mean(20*log10(abs(SpecH(fRange2Lidx:fRange2Uidx,:))),1)';
STER21 = STER2-STER1;
%plot(timeVecRaw,STER1,'g',timeVecRaw,STER2,'m')
plot(t,STER1,'g',t,STER2,'m')
grid on
title('Short-Term Energy')
xlabel('Time (s)')
ylabel('Short-Term Energy (dB)')
set(gca,"Fontsize",FontSize)
ax = gca;
yl = get(ax,'ylim');
yt = get(ax,'ytick');
h0 = get(ax,'children');
hold on
% [ax,h1,h2] = plotyy(ax,0,0,timeVecRaw,STER2-STER1); % plot for STER Michaela
[ax,h1,h2] = plotyy(ax,0,0,t,STER21); % plot for Chiaoscuro Kristen
delete(h1)
%set(ax(1),'linecolor',get(h0,'color'),'ylim',yl,'ytick',yt)
set(h2,'color','k')
%axis([0,timemax -dispRange 5])
ylabel (ax(2), "'Short-Term Energy ratio (dB)'")
legend('STE Range 1', 'STE Range 2', 'STER (R2-R1) (dB)')
set(gca,"Fontsize",FontSize)
hold off


boxh=figure(4);
boxplot(STER21);
name=[name(1:end-4),'_',vowel,'.wav'];
title(strrep(name,'_','\_'));
ylabel('STER')
set(gca,'ylim',[-40 10])
set(gca,'xlim',[0 2])
grid on
set(gca,"Fontsize",FontSize)

savefig(boxh,[figpath,name(1:end-4)])
saveas(boxh,[figpath,name(1:end-4),'.png'])
