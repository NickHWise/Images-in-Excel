e = actxserver('Excel.Application');

eWorkbook = e.Workbooks.Add;
e.Visible = 1;

eSheets = e.ActiveWorkbook.Sheets;
eSheet1 = eSheets.get('Item',1);
eSheet1.Activate

% this makes the cells square
eSheet1.cells.columnwidth=2.43;
eSheet1.cells.rowheight=16.5;

[file, path]=uigetfile;

[I,Imap]=imread(strcat(path,file));

% imshow(I, Imap)

Idims=size(I);

Iwidth=Idims(2);
Iheight=Idims(1);

% Get pixel colours and colour in cell on sheet

for i=1:Iwidth
    for j=1:Iheight
        ref=strcat(number_to_letter(i),num2str(j));
        R=I(j,i,1);
        G=I(j,i,2);
        B=I(j,i,3);
        C = double(R) * 256^0 + double(G) * 256^1 + double(B) * 256^2;
        eSheet1.Range(ref).Interior.Color=C;
    end
end

function [letter] = number_to_letter(n)
    
    n1 = rem(n, 26);
    n2 = floor((n-1)/26);
    
    if n1==0
        n1=n1+26;
    end
    
    let1 = char('A'+n2-1);
    let2 = char('A'+n1-1);
    
    if n2==0 %
        
        letter = let2;
        
    else
    
        letter = strcat(let1,let2);
    
    end
end