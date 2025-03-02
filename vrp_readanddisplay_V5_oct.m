clear all
Fontsize=20;
% requires package statistics for VRP cleaning
pkg load statistics % only needed in octave to load the signal package
pkg load io % only needed in octave to load and save xlsx files

%path='c:\Users\Malte Kob\Downloads\';
%path='c:\Users\Malte Kob\Filr\Meine Dateien\Projekte\LSME\LingWaves\';
%name='lwintern355_20230328_184316.vph';
%path='c:\Users\Malte Kob\Filr\Meine Dateien\Projekte\LSME\LingWaves\';
%path = 'c:\Users\Malte Kob\Downloads\Vokale\';
%path = 'c:\Users\kob\Filr\Meine Dateien\Projekte\TransStimme\LW\lwintern361\04102023\';
path = 'c:\Users\kob\Documents\GitHub\analysistools\';
name='lwintern113_20230221_143534.vph';
%name='lwintern165_20100922_094703.vph';
%name='lwintern165_20160226_161046.vph';
%name = 'lwintern361_20230628_152236.vph';
%name='lwintern165_20230328_210802.vph';
%path='c:\Users\Malte Kob\Filr\Für mich freigegeben\LSS ME\neue Aufnahmen\LingWaves\2023-03-28_S03_396\';
%name='lwintern165_20230328_210802.vph'; % OK
%name='lwintern165_20160226_161046.vph'; % klappt nicht
%name='lwintern165_20100922_094703.vph'; % klappt nicht
%path='c:\Users\Malte Kob\Filr\Für mich freigegeben\LSS ME\neue Aufnahmen\LingWaves\2023-03-28_S02_583\';
%name='lwintern5_20230328_201445.vph';

%path='c:\Users\Malte Kob\Filr\Meine Dateien\Projekte\TransStimme\LW\lwintern361\';
%path='c:\Users\Malte Kob\Filr\Meine Dateien\Projekte\TransStimme\LW\lwintern361\04102023\';
%name='lwintern361_20231004_162627.vph';
%name='lwintern361_20230628_152236.vph';
%path='c:\Users\Malte Kob\Filr\Meine Dateien\Projekte\LSME\LingWaves\';
%name='lwintern113_20230221_143534.vph';

outrange = 90; % average range of values to remove outliers
csvfile=[path, name];
fid = fopen(csvfile);
vrptext = textscan(fid,"%s");
fclose(fid);
vrptext = cell2mat(vrptext);

offset = 61; % start of VRP extreme values (16..19) 2023
%offset = 62; % start of VRP extreme values (16..19) 2010
%offset = 64; % start of VRP extreme values (16..19) 2016

xtract = vrptext{offset};
fmax=str2num(xtract)

xtract = vrptext{offset+5};
fmin=str2num(xtract)

xtract = vrptext{offset+10};
Lmax=str2num(xtract)

xtract = vrptext{offset+15};
Lmin=str2num(xtract)

xtract = vrptext{offset+17};
fRangemin=str2num(xtract)

xtract = vrptext{offset+116};
fRangemax=str2num(xtract)

for idx = 1:100
  counter= idx-1+offset+17;
  freqvec(idx) = str2num(vrptext{counter});
end

counterVRPstart = counter+1;
xtract = vrptext{counterVRPstart};
Lrangemax=str2num(xtract)

counter=counterVRPstart+79*101;
xtract = vrptext{counter};
Lrangemin=str2num(xtract)

clf
for rowidx=1:Lrangemax-Lrangemin+1
    lineidx = counterVRPstart+(rowidx-1)*size(freqvec,2);
    for colidx = 1:size(freqvec,2)
      idx = lineidx+colidx+rowidx-1;
      VRPstr=vrptext{idx};
      VPR(rowidx,colidx) = str2num(VRPstr);
    end
end
Lvec = [Lrangemax:-1:Lrangemin];
[M,c]=contour(freqvec,Lvec,VPR);

xlabel('Frequenz (Hz)')
ylabel('SPL (dB)')
title([strrep(name,'_','\_'), ' - fmax:',num2str(fmax), 'Hz, fmin:',num2str(fmin), 'Hz, Lmax:',num2str(Lmax), 'dB, Lmin:',num2str(Lmin), 'dB']);
grid on
set(gca,"XScale","log")
set(gca,"XMinorGrid","on")
set (gca,"yminorgrid","on");
set(gca,"Fontsize",5)
set(gca,"Xlim",[fRangemin-1 fRangemax+100])
set(gca,"Ylim",[Lrangemin-1 Lrangemax+1])

minSPL=zeros(1,100);
maxSPL=zeros(1,100);
dada=get(c,"zdata");
for Fidx= 1:100
    try
        minSPLidx(Fidx)=max(find(dada(:,Fidx)));
        minSPL(Fidx) = Lvec(minSPLidx(Fidx));
        maxSPLidx(Fidx)=min(find(dada(:,Fidx)));
        maxSPL(Fidx) = Lvec(maxSPLidx(Fidx));
    end
end
minSPLl  = minSPL(find(minSPL));
minFreq = freqvec(find(minSPL));
maxSPLl  = maxSPL(find(maxSPL));
maxFreq = freqvec(find(maxSPL));
line(minFreq,minSPLl,'LineWidth',1,'Color', 'c')
line(maxFreq,maxSPLl,'LineWidth',1,'Color', 'm')

TFmax = isoutlier (maxSPLl, "movmedian", outrange, "SamplePoints", maxFreq);
TFmin = isoutlier (minSPLl, "movmedian", outrange, "SamplePoints", minFreq);
line(maxFreq(~TFmax),maxSPLl(~TFmax),'LineWidth',3,'Color', 'r')
line(minFreq(~TFmin),minSPLl(~TFmin),'LineWidth',3,'Color', 'b')
fname=[name(1:end-4)];
fname=strrep(fname,'.','_');
clear ExcelContent;
%smax = size(maxSPLl(~TFmax),2);
%smin = size(minFreq(~TFmin),2);
%if (smax > smin)
%nzero=zeros(1,smax-smin);
%minF = [minFreq(~TFmin),nzero];
%minL = [minSPLl(~TFmin),nzero];
%maxF = maxFreq(~TFmax);
%maxL = maxSPLl(~TFmax);
%elseif (smax < smin)
%nzero=zeros(1,smin-smax);
%minF = minFreq(~TFmin);
%minL = minSPLl(~TFmin);
%maxF = [maxFreq(~TFmax),nzero];
%maxL = [maxSPLl(~TFmax),nzero];
%else
maxF = maxFreq(~TFmax);
maxL = maxSPLl(~TFmax);
minF = minFreq(~TFmin);
minL = minSPLl(~TFmin);
%end
%ExcelContent= [[maxFreq(~TFmax), minFreq(~TFmin)];[maxSPLl(~TFmax),minSPLl(~TFmin)]];
ExcelContent= [{[path,'VRP_',fname]}];
rstatus = xlswrite ([path,'VRP_',fname], ExcelContent,'A1:A1');
ExcelContent= [{'fsoft'},{'Lsoft'},{'floud'},{'Lloud'}];
rstatus = xlswrite ([path,'VRP_',fname], ExcelContent,'A2:D2');
ExcelContent= [[minF'],[minL']];
rstatus = xlswrite ([path,'VRP_',fname], ExcelContent,'A3:B1000');
ExcelContent= [[maxF'],[maxL']];
rstatus = xlswrite ([path,'VRP_',fname], ExcelContent,'C3:D1000');

ca={"< A1/,A","< A2/A", "< A3/a", "< A4/a'", "< A5/a''", "< A6/a'''"};
cdes={"< Db2/Des", "< Db3/des", "< Db4/des'", "< Db5/des''", "< Db6/des'''"};
%cd={"< D2/D", "< D3/d", "< D4/d'", "< D5/d''", "< D6/d'''"};
cf={"< F2/F", "< F3/f", "< F4/f'", "< F5/f''", "< F6/f'''"};
fa  =55;
fdes=69.2957;
%fd  =73.4162;
ff  =87.3071;
ypos =44;
fontreduction = 9;
for fidx = 0:4
  text(2^fidx*fa, ypos, ca(fidx+1),"Fontsize",Fontsize-fontreduction)
  text(2^fidx*fdes, ypos, cdes(fidx+1),"Fontsize",Fontsize-fontreduction)
  text(2^fidx*ff, ypos, cf(fidx+1),"Fontsize",Fontsize-fontreduction)
  line([2^fidx*fa, 2^fidx*fa],[Lrangemin-1 Lrangemax+1],'LineWidth',1,'LineStyle',':','Color', 'g')
  line([2^fidx*fdes, 2^fidx*fdes],[Lrangemin-1 Lrangemax+1],'LineWidth',1,'LineStyle',':','Color', 'g')
  line([2^fidx*ff, 2^fidx*ff],[Lrangemin-1 Lrangemax+1],'LineWidth',1,'LineStyle',':','Color', 'g')
endfor

set(gca,"Fontsize",Fontsize)
saveas(gcf,[path,name(1:end-4),'_VRP.png'])
