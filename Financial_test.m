clc,clear
filename = 'C:\Users\杨晓乐\Desktop\Financial _Data\example_AP.xlsx';
filename1 = 'C:\Users\杨晓乐\Desktop\Financial _Data\example_back_up.xlsx';

sheet = 1;
xlrange = 'B:C'
[date,year] = xlsread(filename);
[date_orignal,year_orignal] = xlsread(filename1);
% date_1 = {date};

row = 49;
column = 16;

  Sel_name = 3;% the column of information reqired(is the name of company) 
  Sel_investment = 13;% amount of investment (in 13-th column)
  start = 2;
  End = 48;
  Inv_value1  = double(date(:,Sel_investment));
  %%%%%% this circulation is to label defferent kind of companies
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
%%%%%%%%%%%

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
character = 8;
types = 9:12;
Num = date(:,character);
for count_cell = 1:r
    vector_s = cell2mat(vec_s(count_cell));
    vector_e = cell2mat(vec_e(count_cell));
    if isempty(vector_s)
        continue
    else
        for  count_seq= 1:length(vector_s)
            Sel_chara = vector_s(count_seq):vector_e(count_seq);
                if isempty(find(Num(Sel_chara)==2))
                       continue
                elseif (rank(date(Sel_chara,types))>=2)
                       sp=vector_s(count_seq);
                       number(sp,:) = date(sp,1:3);                                      
                else
                    continue
                end    
        end
    end       
end
number_cel = cell(49,16);
[res,indice] = find(number(:,1)>0);
[date_r,date_c] = size(date);
for t =1:length(res)
   for idx_col =1: date_c
      if ((idx_col>=4 && idx_col<=7)||idx_col==14)
              number_cel(res(t),idx_col) = year(res(t),idx_col-3);
      else
               number_cel(res(t),idx_col) = {date(res(t),idx_col)};
      end
   end
end
 xlswrite('C:\Users\杨晓乐\Desktop\Financial _Data\国企和其他企业投资共存.xlsx',number_cel,sheet);