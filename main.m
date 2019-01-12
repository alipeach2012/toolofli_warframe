[template_num,template,~]=xlsread('Ä£°æ.xlsx');
[warehouse_num,warehouse,~]=xlsread('²Ö¿â.xlsx');

itemname={'Õ½¼×','Ç¹','ÊÖÇ¹','Ë«Ç¹', '¹­', 'ÔÓ','È­','È¦'};

for i=1:length(warehouse)
   name=warehouse{i,1};
   part=warehouse{i,2};
   num=warehouse_num(i);
   x=0;
   for j=1:size(template,1)
       if strcmp(template{j,2},name)
           x=j;
           break;
       end
   end
   if x==0
       disp(['please add',name,' to template']);
       continue
   end
       
   xx=x;
   while (xx>1)
       xx=xx-1;
       flag=0;
       for j=1:length(itemname)
           if strcmp(template{xx,1},itemname{j})
               flag=1;
               break;
           end

       end
      if flag==1 
           break;
       end
   end
   
   y=0;
   for j=1:size(template,2)
       if ~isempty(strfind(template{xx,j},part))
           y=j;
           break;
       end
   end
   if y==0
       disp(['please add ',part,' of ' ,name,' to template']);
       continue
   end
     
   template{x,y}=num; 
       
   
end

xlswrite('Çé¿öÍ³¼Æ.xlsx',template);

