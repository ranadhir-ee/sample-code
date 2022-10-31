tic
clear all; close all; clc;

%% Admittance matrix
num = 118;
linedata = linedatas(num);

fb = linedata(:,1);    
tb = linedata(:,2);    
x  = linedata(:,4);    

nbus = max(max(fb),max(tb));
nbranch = length(fb);
Y = zeros(nbus,nbus);

%Off-diagonal elements
 for k=1:nbranch
     Y(fb(k),tb(k)) = -inv(x(k));
     Y(tb(k),fb(k)) = Y(fb(k),tb(k));
 end
 
 %Diagonal elements
  for m =1:nbus
     for n =1:nbranch
         if fb(n) == m
             Y(m,m) = Y(m,m)+inv(x(n));
         elseif tb(n) == m
             Y(m,m) = Y(m,m) + inv(x(n));
         end
     end
    end
 
% Y_Bus Matrix copied to insert perturbation
Y_1p=Y;

Results = [];
Line_max = [];

%%

%list of all lines
lines = [1,2;1,3;2,12;3,5;3,12;4,5;4,11;5,6;5,8;5,11;6,7;7,12;8,9;8,30;9,10;11,12;11,13;12,14;12,16;12,117;13,15;14,15;15,17;15,19;15,33;16,17;17,18;17,30;17,31;17,113;18,19;19,20;19,34;20,21;21,22;22,23;23,24;23,25;23,32;24,70;24,72;25,26;25,27;26,30;27,28;27,32;27,115;28,29;29,31;30,38;31,32;32,113;32,114;33,37;34,36;34,37;34,43;35,36;35,37;37,38;37,39;37,40;38,65;39,40;40,41;40,42;41,42;42,49;42,49;43,44;44,45;45,46;45,49;46,47;46,48;47,49;47,69;48,49;49,50;49,51;49,54;49,54;49,66;49,66;49,69;50,57;51,52;51,58;52,53;53,54;54,55;54,56;54,59;55,56;55,59;56,57;56,58;56,59;56,59;59,60;59,61;59,63;60,61;60,62;61,62;61,64;62,66;62,67;63,64;64,65;65,66;65,68;66,67;68,69;68,81;68,116;69,70;69,75;69,77;70,71;70,74;70,75;71,72;71,73;74,75;75,77;75,118;76,77;76,118;77,78;77,80;77,80;77,82;78,79;79,80;80,96;80,81;80,97;80,98;80,99;82,83;82,96;83,84;83,85;84,85;85,86;85,88;85,89;86,87;88,89;89,90;89,90;89,92;89,92;90,91;91,92;92,93;92,94;92,100;92,102;93,94;94,95;94,96;94,100;95,96;96,97;98,100;99,100;100,101;100,103;100,104;100,106;101,102;103,104;103,105;103,110;104,105;105,106;105,107;105,108;106,107;108,109;109,110;110,111;110,112;114,115];

%Area G1
A_G1 = [77,80;77,80;77,82;78,79;79,80;80,96;80,81;80,97;80,98;80,99;82,83;82,96;83,84;83,85;84,85;85,86;85,88;85,89;86,87;88,89;89,90;89,90;89,92;89,92;90,91;91,92;92,93;92,94;92,100;92,102;93,94;94,95;94,96;94,100;95,96;96,97;98,100;99,100;100,101;100,103;100,104;100,106;101,102;103,104;103,105;103,110;104,105;105,106;105,107;105,108;106,107;108,109;109,110;110,111;110,112];
%Area G2
A_G2 = [38,65;42,49;42,49;43,44;44,45;45,46;45,49;46,47;46,48;47,49;47,69;48,49;49,50;49,51;49,54;49,54;49,66;49,66;49,69;50,57;51,52;51,58;52,53;53,54;54,55;54,56;54,59;55,56;55,59;56,57;56,58;56,59;56,59;59,60;59,61;59,63;60,61;60,62;61,62;61,64;62,66;62,67;63,64;64,65;65,66;65,68];
%Area G3
A_G3 = [1,2;1,3;2,12;3,5;3,12;4,5;4,11;5,6;5,8;5,11;6,7;7,12;8,9;8,30;9,10;11,12;11,13;12,14;12,16;12,117;13,15;14,15;15,17;15,19;15,33;16,17;17,18;17,30;17,31;17,113;18,19;19,20;19,34;20,21;21,22;22,23;23,24;23,25;23,32;24,70;24,72;25,26;25,27;26,30;27,28;27,32;27,115;28,29;29,31;30,38;31,32;32,113;32,114;33,37;34,36;34,37;34,43;35,36;35,37;37,38;37,39;37,40;38,65;39,40;40,41;40,42;41,42;42,49;43,44;114,115];
%Area G4
A_G4 = [68,69;68,81;68,116;69,70;69,75;69,77;70,71;70,74;70,75;71,72;71,73;74,75;75,77;75,118;76,77;76,118;77,78];
%Selected lines
L_c = [89,90;89,90;37,38;89,92;89,92;26,30;38,65;17,30;49,66;49,66;59,63;63,64;42,49;42,49;8,30;23,25;66,67;69,70;64,65;49,51;77,80;77,80];


%combination of 4 lines at k=4 stored in list L
a=1;
List_L=[];
for r=1:length(L_c)
    for s=(r+1):length(L_c)
         for t = (s+1):length(L_c)
             for u = (t+1):length(L_c)
                if L_c(r,:) == L_c(s,:) == L_c(t,:) == L_c(u,:)
                     break
                else
                   List_L(a,:) = [L_c(r,:) L_c(s,:) L_c(t,:) L_c(u,:)];
                   a=a+1;
          end
          end
         end
    end
end 

%%modify Y-bus matrix, constraint matrix formation, LP solution for each
%%k-line in the list L

for k_line = 1:length(List_L)
    Y_1p=Y;
    
    % Input Bus number 1 to 4
     node_i = List_L(k_line,1);
     node_j = List_L(k_line,2);
     node_i1 = List_L(k_line,3);
     node_j1 = List_L(k_line,4);
     node_i2 = List_L(k_line,5);
     node_j2 = List_L(k_line,6);
     node_i3 = List_L(k_line,7);
     node_j3 = List_L(k_line,8);
     % Input percentage of perturbation 
     per = 100; 

     % Calculates the perturbation on admittance of the lines 
     per = per/100;
     del = per*Y_1p(node_i,node_j);
     del1 = per*Y_1p(node_i1,node_j1);
     del2 = per*Y_1p(node_i2,node_j2);
     del3 = per*Y_1p(node_i3,node_j3);

     % Incorporates the perturbation with the Y_bus elements for line 1

     Y_1p(node_i,node_j)=Y_1p(node_i,node_j)-del;
     Y_1p(node_j,node_i)=Y_1p(node_j,node_i)-del;
     Y_1p(node_i,node_i)=Y_1p(node_i,node_i)+del;
     Y_1p(node_j,node_j)=Y_1p(node_j,node_j)+del;

     % Incorporates the perturbation with the Y_bus elements for line 2

     Y_1p(node_i1,node_j1)=Y_1p(node_i1,node_j1)-del1;
     Y_1p(node_j1,node_i1)=Y_1p(node_j1,node_i1)-del1;
     Y_1p(node_i1,node_i1)=Y_1p(node_i1,node_i1)+del1;
     Y_1p(node_j1,node_j1)=Y_1p(node_j1,node_j1)+del1;
     
     % Incorporates the perturbation with the Y_bus elements for line 3

     Y_1p(node_i2,node_j2)=Y_1p(node_i2,node_j2)-del2;
     Y_1p(node_j2,node_i2)=Y_1p(node_j2,node_i2)-del2;
     Y_1p(node_i2,node_i2)=Y_1p(node_i2,node_i2)+del2;
     Y_1p(node_j2,node_j2)=Y_1p(node_j2,node_j2)+del2;
     
     % Incorporates the perturbation with the Y_bus elements for line 4

     Y_1p(node_i3,node_j3)=Y_1p(node_i3,node_j3)-del3;
     Y_1p(node_j3,node_i3)=Y_1p(node_j3,node_i3)-del3;
     Y_1p(node_i3,node_i3)=Y_1p(node_i3,node_i3)+del3;
     Y_1p(node_j3,node_j3)=Y_1p(node_j3,node_j3)+del3;
    
%% Constraint matrix formation
lines = [1,2;1,3;2,12;3,5;3,12;4,5;4,11;5,6;5,8;5,11;6,7;7,12;8,9;8,30;9,10;11,12;11,13;12,14;12,16;12,117;13,15;14,15;15,17;15,19;15,33;16,17;17,18;17,30;17,31;17,113;18,19;19,20;19,34;20,21;21,22;22,23;23,24;23,25;23,32;24,70;24,72;25,26;25,27;26,30;27,28;27,32;27,115;28,29;29,31;30,38;31,32;32,113;32,114;33,37;34,36;34,37;34,43;35,36;35,37;37,38;37,39;37,40;38,65;39,40;40,41;40,42;41,42;42,49;42,49;43,44;44,45;45,46;45,49;46,47;46,48;47,49;47,69;48,49;49,50;49,51;49,54;49,54;49,66;49,66;49,69;50,57;51,52;51,58;52,53;53,54;54,55;54,56;54,59;55,56;55,59;56,57;56,58;56,59;56,59;59,60;59,61;59,63;60,61;60,62;61,62;61,64;62,66;62,67;63,64;64,65;65,66;65,68;66,67;68,69;68,81;68,116;69,70;69,75;69,77;70,71;70,74;70,75;71,72;71,73;74,75;75,77;75,118;76,77;76,118;77,78;77,80;77,80;77,82;78,79;79,80;80,96;80,81;80,97;80,98;80,99;82,83;82,96;83,84;83,85;84,85;85,86;85,88;85,89;86,87;88,89;89,90;89,90;89,92;89,92;90,91;91,92;92,93;92,94;92,100;92,102;93,94;94,95;94,96;94,100;95,96;96,97;98,100;99,100;100,101;100,103;100,104;100,106;101,102;103,104;103,105;103,110;104,105;105,106;105,107;105,108;106,107;108,109;109,110;110,111;110,112;114,115];

F2=zeros(nbranch,nbus-1);

 for k=1:nbranch
     i=lines(k,1);
     j=lines(k,2);
     if i~=1
         F2(k,i-1)=-Y_1p(i,j);
         F2(k,j-1)=Y_1p(i,j);
     else
         F2(k,j-1)=Y_1p(i,j);
     end
 end
 
F1 = [10.01;23.5849;zeros(184,1)];
F = [F1 F2];
C = F';
C_matrix = zeros(nbus,nbranch);

for p = 1:nbus
    for q = 1:nbranch
      if C(p,q)>0
           C_matrix(p,q) = -1;
          else if C(p,q) == 0
           C_matrix(p,q) = 0;
          else if C(p,q) < 0
           C_matrix(p,q) = 1;      
           end
           end
      end
    end
end

%% Linear Programing 
% cost function
fun = [ones(1,nbus) zeros(1,nbranch)]';
% bus injection
Pi = [-51;-20;-39;-39;0;-52;-19;-28;0;450;-70;38;-34;-14;-90;-25;-11;-60;-45;-18;-14;-10;-7;-13;220;314;-71;-17;-24;0;-36;-59;-23;-59;-33;-31;0;0;-27;-66;-37;-96;-18;-16;-53;-9;-34;-20;117;-17;-17;-18;-23;-65;-63;-84;-12;-12;-122;-78;160;-77;0;0;391;353;-28;0;381;-66;0;-12;-6;-68;-47;-68;-61;-71;-39;347;0;-54;-20;-11;-24;-21;4;-48;607;-163;-10;-65;-12;-30;-42;-38;-15;-34;-42;215;-22;-5;17;-38;-31;-43;-50;-2;-8;-39;36;-68;-6;-8;-22;-184;-20;-33];

% Automation to generate P_min and P_max vectors in the constraints of the MINLP
for h=1:length(Pi)
if Pi(h)>0
    P_min(h)=0;
else if Pi(h)<=0
    P_min(h)=Pi(h);
end
end    
end
P_min=P_min';

for h=1:length(Pi)
if Pi(h)>0
    P_max(h)=Pi(h);
else if Pi(h)<=0
    P_max(h)=0;
end
end    
end

P_max=P_max';

%Line rating
Pij_max = [90;75;110;100;60;275;235;365;500;340;315;300;450;465;450;75;230;200;200;50;200;200;250;220;55;220;250;450;250;50;185;230;135;250;260;270;235;450;335;200;165;335;185;350;50;110;100;50;50;300;275;50;70;50;65;110;155;50;65;400;100;255;400;70;50;160;50;100;300;175;190;320;235;350;50;65;385;50;60;360;50;50;355;355;85;50;330;60;320;325;50;100;400;220;275;50;70;0;185;275;175;700;185;355;250;330;325;430;700;360;110;350;460;150;200;185;440;430;120;150;125;250;150;50;70;185;130;105;100;110;100;400;150;165;200;75;200;115;200;50;230;170;75;250;75;50;100;250;50;135;400;400;200;560;225;215;100;125;300;50;85;50;110;75;50;100;235;50;50;230;75;265;50;50;65;200;50;125;175;115;220;120;125;50;70;80];

%Inequality matrix
Aineq = [-eye(nbus,nbus) C_matrix; -eye(nbus,nbus) -C_matrix; zeros(nbus,nbus) -C_matrix; zeros(nbus,nbus) C_matrix; ...
         zeros(nbranch,nbus) -eye(nbranch,nbranch); zeros(nbranch,nbus) eye(nbranch,nbranch)];
     

% Automation to generate r.h.s of Inequality constraint Ax \leq b
rem=find(lines(:,1)==node_i & lines(:,2)==node_j);
rem1=find(lines(:,1)==node_i1 & lines(:,2)==node_j1);
rem2=find(lines(:,1)==node_i2 & lines(:,2)==node_j2);
rem3=find(lines(:,1)==node_i3 & lines(:,2)==node_j3);

%x_in = 1;
x_out = 0;

temp = 0;
temp1 = 0;
temp2 = 0;
temp3 = 0;

temp = x_out*Pij_max(rem);
temp1 = x_out*Pij_max(rem1);
temp2 = x_out*Pij_max(rem2);
temp3 = x_out*Pij_max(rem3);

for b = 1:nbranch
      if b==rem
          Line_max(b) = temp;
      else if b==rem1
          Line_max(b) = temp1; 
      else if b==rem2
          Line_max(b) = temp2; 
      else if b==rem3
          Line_max(b) = temp3; 
      else if b~=rem 
          Line_max(b) = Pij_max(b);
      else if b~=rem1
          Line_max(b) = Pij_max(b);
      else if b~=rem2
          Line_max(b) = Pij_max(b);
      else if b~=rem3
          Line_max(b) = Pij_max(b);
         end
         end
         end
         end
         end
         end
       end
     end
end

Pmi = [];
Pma = [];

for g = 1:nbus
   Pmi = [Pmi;Pi(g)-P_min(g)];
   Pma = [Pma;P_max(g)-Pi(g)];
end

%r.h.s of Inequality matrix
bineq =  [-(Pi);Pi;Pmi;Pma;Line_max';Line_max'];

%optimization solver for linear programing 
[Var,fval] = linprog(fun,Aineq,bineq,[],[],[],[])


% Get the power flow on the lines from the solution of the LP
g = 119;
Post_Flow=[];
for h=1:length(lines)
    Post_Flow = [Post_Flow;[lines(h,1) lines(h,2) Var(g)]];
    g=g+1;
end

 Post_Flow = [Post_Flow Line_max'];  
 %% 
 
 w_i=1;

New_Flow = zeros(nbus, nbus);

for i = 1:nbranch
    New_Flow(Post_Flow(i,1),Post_Flow(i,2)) = Post_Flow(i,3);
    New_Flow(Post_Flow(i,2),Post_Flow(i,1)) = -Post_Flow(i,3);
end


% Identify critical k-line contingencies and store them in the variable
% Results
cost_value = 0;
for j = 1:nbus
  cost_value = cost_value + w_i*(abs(Pi(j) - sum(New_Flow(j,:))));     
end

%cost_value = [cost_value;cost_valuend];
%Results = [Results;[L(k_line,:) cost_value]];
    
if cost_value >= 3000
Results = [Results;[List_L(k_line,:) cost_value]];
end


Line_max = [];    
p = p + 1;  

end

toc