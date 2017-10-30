FileListAdverse = dir('\\arxhz2\Documents\Risk Analytics\Data Management\Macroeconomic Data\JUL 2017\FRB ADVERSE\Flat Files');
FileListBaseline = dir('\\arxhz2\Documents\Risk Analytics\Data Management\Macroeconomic Data\JUL 2017\FRB BASELINE\Flat Files');
FileListSevere = dir('\\arxhz2\Documents\Risk Analytics\Data Management\Macroeconomic Data\JUL 2017\FRB SEVERELY ADVERSE\Flat Files');

cd('\\arxhz2\Documents\Risk Analytics\Data Management\Macroeconomic Data\JUL 2017\FRB ADVERSE\Flat Files');
Adverse = xlsread(FileListAdverse(3).name);
Adverse(isnan(Adverse)) = 0;
cd('\\arxhz2\Documents\Risk Analytics\Data Management\Macroeconomic Data\JUL 2017\FRB BASELINE\Flat Files');
Baseline = xlsread(FileListBaseline(3).name);
Baseline(isnan(Baseline)) = 0;
cd('\\arxhz2\Documents\Risk Analytics\Data Management\Macroeconomic Data\JUL 2017\FRB SEVERELY ADVERSE\Flat Files');
Severe = xlsread(FileListSevere(3).name);
Severe(isnan(Severe)) = 0;

Time = Adverse(:,1);
TimeCut = find(Time == 2017)+2;

BaseAdverse= Baseline - Adverse;
Idx = find(sum(BaseAdverse(:,2:end),2)~=0);
results{1,1}=FileListBaseline(3).name;
if (Idx(1)<TimeCut)
    results{1,2} = 'Not_statring_from_2017Q3';
else
    results{1,2} = 'statring_from_2017Q3';
end

AdverseSevere= Adverse - Severe;
Idx = find(sum(AdverseSevere(:,2:end),2)~=0);
results{2,1}=FileListAdverse(3).name;
if (Idx(1)<TimeCut)
    results{2,2} = 'Not_statring_from_2017Q3';
else
    results{2,2} = 'statring_from_2017Q3';
end

SevereBaseline = Severe - Baseline;
Idx = find(sum(SevereBaseline(:,2:end),2)~=0);
results{3,1}=FileListSevere(3).name;
if (Idx(1)<TimeCut)
    results{3,2} = 'Not_statring_from_2017Q3';
else
    results{3,2} = 'statring_from_2017Q3';
end

results = results';
fid = fopen('check_scenario.txt','w');
fprintf(fid,'%s %s\n',results{:});
fclose(fid);

%%
excelObject = actxserver('Excel.Application');
FileListAdverse = dir('Z:\Data Management\Macroeconomic Data\JUL 2017\FRB ADVERSE\Flat Files')
FileListBASELINE = dir('Z:\Data Management\Macroeconomic Data\JUL 2017\FRB BASELINE\Flat Files')
FileListSevere = dir('Z:\Data Management\Macroeconomic Data\JUL 2017\FRB SEVERELY ADVERSE\Flat Files')

xlsFiles = FileListAdverse;
for k = 3:length(xlsFiles)
  baseFileName = xlsFiles(k).name;
  fprintf(1, 'Now reading %s\n', baseFileName);
  excelWorkbook = excelObject.workbooks.Open(strcat('Z:\Data Management\Macroeconomic Data\JUL 2017\FRB ADVERSE\Flat Files',baseFileName));
  worksheets = excelObject.sheets;
  numberOfSheets = worksheets.Count;
  for sheetIndex  = 1 : numberOfSheets 
    % Do whatever you want to do.
  end
end
excelObject.Quit;    