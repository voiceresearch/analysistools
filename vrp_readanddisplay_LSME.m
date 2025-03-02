clear all
Fontsize=7;
% requires package statistics for VRP cleaning
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

%path = 'c:\Users\kob\Documents\GitHub\analysistools\';
path = 'c:\temp\LSME\VRP\';

cd(path);
partlist = GetSubDirsFirstLevelOnly(pwd)
clf
%hold on
for partidx = 1:size(partlist,2)
  disp('_________________________________')
  actpart = char(partlist(partidx))
  cd(actpart)
  figpath=path;
  reclist = dir('*.vph');
  for recidx = 1:size(reclist,1)
    disp('_________________________________')
    fname_org = reclist(recidx).name;
    fname=[fname_org(1:end-4)];
    fname=strrep(fname,'.','_')
    fnamestr=strrep(fname,'_','\_');
    fnamecellarray=cellstr(fname);


    outrange = 90; % average range of values to remove outliers
    csvfile=[path, fname_org];
    fid = fopen(fname_org);
    vrptext = textscan(fid,"%s");
    fclose(fid);
    vrptext = cell2mat(vrptext);

    switch (actpart)
      case {"A" "B"}
        offset = 62;
      case {"C"}
        offset = 61;
      otherwise
         error ("invalid value");
    endswitch
    %offset = 61; % start of VRP extreme values (16..19) 2023
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

    %clf
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
    title([strrep(fname,'_','\_'), ' - fmax:',num2str(fmax), 'Hz, fmin:',num2str(fmin), 'Hz, Lmax:',num2str(Lmax), 'dB, Lmin:',num2str(Lmin), 'dB']);
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
    line(maxFreq(~TFmax),maxSPLl(~TFmax),'LineWidth',2,'Color', 'r')
    line(minFreq(~TFmin),minSPLl(~TFmin),'LineWidth',2,'Color', 'b')
%    fname=[name(1:end-4)];
%    fname=strrep(fname,'.','_');
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
    ExcelContent= [{[actpart,'\',fname]}];
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
    fontreduction = 2;
    for fidx = 0:4
      text(2^fidx*fa, ypos, ca(fidx+1),"Fontsize",Fontsize-fontreduction)
      text(2^fidx*fdes, ypos, cdes(fidx+1),"Fontsize",Fontsize-fontreduction)
      text(2^fidx*ff, ypos, cf(fidx+1),"Fontsize",Fontsize-fontreduction)
      line([2^fidx*fa, 2^fidx*fa],[Lrangemin-1 Lrangemax+1],'LineWidth',1,'LineStyle',':','Color', 'g')
      line([2^fidx*fdes, 2^fidx*fdes],[Lrangemin-1 Lrangemax+1],'LineWidth',1,'LineStyle',':','Color', 'g')
      line([2^fidx*ff, 2^fidx*ff],[Lrangemin-1 Lrangemax+1],'LineWidth',1,'LineStyle',':','Color', 'g')
    endfor

    set(gca,"Fontsize",Fontsize)
    saveas(gcf,[path,fname(1:end-4),'_VRP.png'])
  endfor
  cd ..
endfor
%hold off
