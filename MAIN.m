close all; warning off; clc 	% Clear all
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% CONFIGURATIONS
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

% Choose which case(s) will be simulated
Faults = [1,2,3,4,5]; % 1=Normal, 2=ShortCircuit, 3=Degradation, 4=OpenCircuit, 5=Shadow

% Set Fault's Configuration
SCR = 10E-9;    % Set the Short Circuit Resistance (Should be very low [10E-9 for instance])
DGR = 20;       % Set the Degradation Resistance
SHM = 0.5;      % Set the Irradiance multiplier for the Shadow Fault  (Number between 0 [Fully shadded] and 1 [No shade])

% Set the Dataset
File = 'dataset.xlsx';  % Set the Dataset file
MIr = 50;               % Set the minimum irradiance to be used
Columns = [1, 2];       % Columns where the Temperature and Irradiance, respectively, can be found in the dataset's file   

% Set the Simulation's Configs
FS = 1;             % Set Sampling Rate (FS=2 -> Every 1 second in simulation (Simulation time, not 'real' time) = 2 lines of dataset)
% FS = min(FS,1);   % Comment this line if you don't want the FS limmiter
% I don't recommend FS > 1, unless your irradiance
% and temperature changes very little over time. Still recommend FS = 1
% 
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% MAIN
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Fault_Name = {'Normal', 'Short', 'Degrad', 'OpenC', 'Shadow'}; % Set Faults names

% Prepare Dataset
Table=xlsread(File);    % Reads the Dataset
NewTable=[];
for idx = 1:size(Table,1)
    if Table(idx, Columns(2)) > MIr
        NewTable=[NewTable ;[Table(idx,Columns(1)) Table(idx,Columns(2))]];
    end
end

G = NewTable(:, Columns(2));
T = NewTable(:, Columns(1));

% Set Fault's Config File
save=fopen('Config.txt','w'); % Open PSim's Simulation File
fprintf(save,'SCR = %g%+04d;',SCR);
fprintf(save,'\r\n');
fprintf(save,'DGR = %f;',DGR);
fprintf(save,'\r\n');
fprintf(save,'SHM = %f;',SHM);
fprintf(save,'\r\n');
fprintf(save,'FS = %f;',FS);
fclose(save);

clear save
% Run Simulations 
SG = [G];
ST = [T];
for idx = 1:length(Faults)
    PATH=sprintf('%s\\PS_%s.psimsch',pwd,Fault_Name{Faults(idx)}); 
    sim('Simula')
    VDC1 = VDC1(2:end); VDC2 = VDC2(2:end); IDC1 = IDC1(2:end); 
    IDC2 = IDC2(2:end); PAC = PAC(2:end);
    save(strcat('Result_',Fault_Name{idx}), 'VDC1', 'VDC2', 'IDC1', 'IDC2', 'PAC', 'G', 'T')  
    sprintf('Case: %s -> finished', Fault_Name{idx})    
end    
