%%问题三解答
clc,clear
filename = 'C:\Users\杨晓乐\Desktop\Financial _Data\example_AP.xlsx';
filename1 = 'C:\Users\杨晓乐\Desktop\Financial _Data\example_back_up.xlsx';

sheet = 1;
xlrange = 'B:C';
[date,year] = xlsread(filename);
[date_orignal,year_orignal] = xlsread(filename1);

row = 49;
column = 16;

  Sel_name = 3;
  Sel_investment = 13;
 
  start = 2;
  End = 48;
  Inv_value1  = double(date(:,Sel_investment));
 for i=start:End
     if (date(i-1,Sel_name) == date(i,Sel_name) && date(i+1,Sel_name)==date(i,Sel_name))   
                        index(i)=1;
    elseif(date(i-1,Sel_name) == date(i,Sel_name) && date(i+1,Sel_name)~=date(i,Sel_name)) 
                        index(i)=3;
    elseif(date(i-1,Sel_name) ~= date(i,Sel_name) && date(i+1,Sel_name)==date(i,Sel_name)) 
                        index(i)=2;
    elseif(date(i-1,Sel_name) ~= date(i,Sel_name) && date(i+1,Sel_name)~=date(i,Sel_name)) 
                        index(i)=4;
    end      
 end

if (date(End+1,Sel_name)~=date(End,Sel_name))
          index(End+1) = 2;
else
          index(End+1) = 3;
end
if (date(start-1,Sel_name)~=date(start,Sel_name))
          index(start-1) = 3;
else
          index(start-1) = 2;
end

index = index';
idx_4 = find(index==4);
index(idx_4) = 0;
Inv = zeros(49,1);
Inv(idx_4) = Inv_value1(idx_4);

set = 10;
vec_s = cell(set,length(set)) ;
vec_e = cell(set,length(set)) ;
for p_len=0:1:set
    add = ones(1,p_len);
    if isempty(add)
        pat = [2 3];
    else
    pat = [2 add 3];
    end
       if isempty(findpattern(index,pat))
           break
       else
           vec_s(p_len+2,:) = {findpattern(index,pat)};
           vec_e(p_len+2,:) = {findpattern(index,pat)+p_len+1}; 
      
       end
end
[r,c] = size(vec_s);
for count_cell = 1:r
    vector_s = cell2mat(vec_s(count_cell));
    vector_e = cell2mat(vec_e(count_cell));
    if isempty(vector_s)
        continue
    else
        for  count_seq= 1:length(vector_s)
            temp = Inv_value1( vector_s(count_seq):vector_e(count_seq));
            Inv(vector_s(count_seq):vector_e(count_seq)) = max(temp);
            [val,offset] = max(temp);
            Num(vector_s(count_seq)+offset-1)=date(vector_s(count_seq)+offset-1,Sel_name);           
       
        end
    end
end
Num = Num';

max_cel = cell(49,16);
[res,indice] = find(Num(:,1)>0);
[date_r,date_c] = size(date);
for t =1:length(res)
   for idx_col =1: date_c
      if ((idx_col>=4 && idx_col<=7)||idx_col==14)
              max_cel(res(t),idx_col) = year(res(t),idx_col-3);
      else
              max_cel(res(t),idx_col) = {date(res(t),idx_col)};
      end
   end
end
xlswrite('C:\Users\杨晓乐\Desktop\Financial _Data\单个被投资企业投资金额最高值.xlsx',max_cel,sheet);