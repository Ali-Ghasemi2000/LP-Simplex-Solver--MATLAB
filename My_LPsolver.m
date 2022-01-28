function My_LPsolver

% This MATLAB Function aims to solve LPs showing the simplex table step-by-step. 
% The user inserts the column number which becomes a basic variable and the 
% row number which goes out and becomes non-basic variable. 

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
              By MIR ALI GHASEMI
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Excel File MUST be is the Following Format:
% g: greater than
% l: less than
% e: equal to

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% {min/max} {x1}      {x2} ...     {xn}      {state} {rhs}
% {c1}      {a1_1}    {a1_2}       {a1_n}    {g/l/e} {b1}
% {c1}      {a2_1}    {a2_2}       {a2_n}    {g/l/e} {b2}
% .
% .
% .
% {cm}      {am_1}    {am_2}       {am_n}    {g/l/e} {bm}

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Get Excel File Name
file_in1 = input("Enter Excel File Name (WITHOUT '.xlsx') : \n", 's');
file_in2 = file_in1+".xlsx";

% Convert Table to Cell
raw = table2cell(readtable(file_in2, 'ReadVariableNames',false));
[nr, nc] = size(raw);

% Minimization OR Maximization
minormax = raw{1,1};

% Read Constraints : >=     <=      =
sts = cell2mat(raw(3:end, nc-1));

% Num of Vars (x)
numx = nc-3;

% GREATER THAN (>)  
numg = sum(sts=='g');

% LESS THAN (<)
numl = sum(sts=='l');

% EQUAL TO (=)
nume = sum(sts=='e');

% Num of Slack Vars (s)
nums = numg + numl;

% Num of Artificial Vars (r)
numr = numg + nume;

% Def Vars
syms 'x%d' [1, numx]
syms 's%d' [1, nr-2]
ST = true(nr-2, 1);
if numr > 0
    syms 'r%d' [1, nr-2]
    RT = true(nr-2, 1);
    syms M
    R = sym(zeros(nr-1, numr));

end

% Cast X matrix
X = sym(raw(2:end, 2:end-2));
X(1,:) = X(1,:)*-1;

% Cast Right Hand Side (RHS)
RHS = sym(raw(2:end, end));

% Pre-Def S
S = sym(zeros(nr-1, nums));

% Pre-Def Left Hand Side (LHS)
LHS = sym(zeros(nr-2,1));
LHS(1) = sym('z');

% Fill LHS (Constraints), R & S 

Rrows = [];
for i=1:nr-2
    if sts(i)=='l'
        LHS(i+1) = s(i);
        S(i+1,i)= sym(+1);
        RT(i) = false;
    elseif  sts(i)=='g'
        LHS(i+1) = r(i);
        R(i+1,i)= sym(+1);
        Rrows = [Rrows i+2];
        S(i+1,i)= sym(-1);
    else
        LHS(i+1) = r(i);
        R(i+1,i)= sym(+1);
        Rrows = [Rrows i+2];
        ST(i)= false;
    end
end
S = S(:, any(S));

% if any R => Add
% SubT : Inner Matrix of Table
if numr>0
    R = R(:, any(R));
    if minormax=='min'
        R(1,:)= repmat(-M,1,numr); 
    else
        R(1,:)= repmat(M,1,numr);
    end
    subT = [LHS X S R RHS];
else
    subT = [LHS X S RHS]; 
end
if numr>0
        r = r(RT);
end

s = s(ST);

% First Row (Top Row)
Frow = sym(zeros(1, numx+ nums+ numr+2));
Frow(1) = minormax;
Frow(end) = sym('rhs');
Frow(2:numx+1) = x; 
Frow(numx+2: end-1-numr) = s;

% If any R => Add
if numr>0
    Frow(Frow==0) = r;
end

% Counter of Tables in Solution Excel File
xlrng = 1;

% Starting Point in Solution Excel File
xlArng = sprintf("A%d", xlrng);

% Name of Solution Excel File
file_out = file_in1 + "SOL.xlsx";

% First Simplex Table 
T0 = [Frow;subT]

% Write to Solution Excel File 
my_Excel_writer(file_out, T0, xlrng, nr)

% Add ONE to xlrng => Next Table
xlrng = xlrng +1;

% If any R => Pivoting Needed
if numr>0    
    if minormax=='min'
        T0(2,2:end) = T0(2,2:end) + sum(M*T0(Rrows,2:end))
    else
        T0(2,2:end) = T0(2,2:end) + sum(-M*T0(Rrows,2:end))
    end
    my_Excel_writer(file_out, T0, xlrng, nr)
    xlrng = xlrng +1;
end

% Invoke my_pivoting => Step-by-Step Solution
T0 = my_pivoting(T0, file_out, xlrng);

end



function T = my_pivoting(t, FNAME, xrng)

% This Function Gets a Simplex Table and Solves Step-By-Step

% First Row Vars
ROW = t(1,:);

% First Column Vars (Basic Vars)
COL = t(2:end, 1);

% Inner Matrix (Coefs)
t = t(2:end, 2:end);

% Ask User to Choose a Col to Go in 
col = input("number of COLUMN going IN: ") -1;

% Keep Replacing Basic Vars => Reach Optimum Answer
while true
    
    % Ratio Vector (Minimum Positive)
    vpa(t(2:end, end)./ (t(2:end, col)))
    row = input("number of ROW going OUT : ")+1;
    
    % Chosen Row
    rout = t(row, :);
    
    % Pivot Element
    p = t(row, col);
    
    % Coef to Pivot the Col
    coef = -t(:,col)./ p;
    
    % Pivot the Col
    [R, C] = meshgrid(rout, coef);
    t = t + R .* C;
    
    % Divide Leaving Row By Pivot Element
    t(row,:) = rout./ p;
    
    % Bring the Chosen Var in (Make Basic)
    COL(row) = ROW(col+1);
    
    % New Simplex Table 
    T = [ROW; [COL t]]
    
    % Write New Table to SOL File
    my_Excel_writer(FNAME,T, xrng, (2+size(t,1)))
    
    % Add ONE to xrng (New Table)
    xrng = xrng + 1;
    
    % Ask for New Col (going out)
    % If Optimum => ZERO
    col = input("number of COLUMN going IN: (if Optimum, insert '0') \n")-1;
    
    % IF Optimum, END While Loop
    if col ==-1 
        break
    end
end

% Final SOL Table 
T = [ROW; [COL t]];

end



function my_Excel_writer(Fname, T, xlrng, nr)
    % Fname: File Name to Write on
    % T    : Table
    % xlrng: Num of Starting Point in Excel
    % nr   : Num Of Rows of the Table
    if xlrng==1
        % Sheet 1   Cell A1
        xlswrite(Fname,string(vpa(T)), 1, 'A1')
    else
        % Compose Cell Based on xlrng
        xlRange = sprintf("A%d", (xlrng-1)*(nr+2));
        xlswrite(Fname,string(vpa(T,2)), 1, xlRange)
    end
end
