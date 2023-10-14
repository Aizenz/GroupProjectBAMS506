# %%
from gurobipy import (GRB,Model,quicksum)
import pandas as pd
import numpy as np
# %%
class EzGurobi:
    # variables list
    data = None 
    obj = None
    obj_coeff = []
    n_variables = []
    variables_name = []
    variables_type = []
    variables_lb = []
    variables_up = []
    sensitive_analysis = None 
    cons_df = None 
    def ezload(self, filename):
        self.data = pd.read_excel(filename,header=None)
        # Step 2: Read and formating all the data into pandas and deal with wrong input data

        # read objective
        obj = self.data.iloc[0,1]
        if not obj in ['max','min']:
            raise NameError('Objective cannot be recognized')
        else:obj= np.where(obj=='max',1,0)


        # read obj coefficient and the number of variables
        self.obj_coeff = pd.Series(data = self.data.iloc[2,1:],name = self.data.iloc[2,0],dtype=float).dropna().reset_index(drop=True)
        self.n_variables = len(self.obj_coeff)

        # read variable names
        self.variable_name = pd.Series(data = self.data.iloc[1,1:],name = self.data.iloc[1,0]).dropna().reset_index(drop=True).to_list()
        if len(self.variable_name) == 0: 
            # automatically generated variables when names are not specified
            self.variable_name.append(['x{}'.format(i)for i in range(self.n_variables)])

        # read variable types
        self.variable_type = pd.Series(data=self.data.iloc[3,1:],name=self.data.iloc[3,0]).dropna().reset_index(drop=True).to_list()
        if len(self.variable_type) == 0:
            self.variable_type = np.repeat('C',self.n_variables)
        elif not np.count_nonzero(self.variable_type.isin(['C','I','B'])):
            raise NameError('Variable type shoud be C I B') 

        # read variable lower bound
        self.variable_lb = self.data.iloc[4,1:].dropna().to_list()
        if "+inf" in self.variable_lb:
            raise NameError('Variable lower bound cannot be +inf') 
        elif len(self.variable_lb)==0:
            self.variable_lb = np.repeat(0,self.n_variables)
        self.variable_lb = [-GRB.INFINITY if i=="-inf" else float(i) for i in self.variable_lb]

        # read upper bound

        self.variable_ub = self.data.iloc[5,1:].dropna().to_list()
        if len(self.variable_ub)==0:
            self.variable_ub = [GRB.INFINITY for i in range(self.n_variables)]
        elif "-inf" in self.variable_ub:
            raise NameError('Variable upper bound cannot be -inf') 
        self.variable_ub = [GRB.INFINITY if i =='+inf' else float(i) for i in self.variable_ub]

        # whether to run sensitivity analysis
        run_sa = self.data.iloc[6,1] 
        
        # Step 3: read constraints
        ## find constraint column index
        const_typ_cind = self.data.loc[7,self.data.iloc[7].isin(['constraint type'])].index[0]

        if const_typ_cind-1 != self.n_variables:
            raise NameError('Number of variables does not match the constraints')

        ## find last constraint row index
        last_r_consraints = len(self.data[const_typ_cind]) 

        ## read constraints as dataframe
        self.cons_df = pd.DataFrame(self.data.loc[8:,1:].fillna(0)).reset_index(drop=True)
        self.cons_df.columns = np.append(self.variable_name,['constraint type','RHS values'])
    
    def ezrun(self,obj='min', run_sa = False):
        
        #if self.data == None: raise ValueError("data is required,load data first")
        m = Model()
        x=m.addVars(self.n_variables)
        # set types, lb, ub of variables
        for i in range(self.n_variables):
            x[i].setAttr('VarNAME', self.variable_name[i])
            x[i].setAttr('vType', self.variable_type[i])
            x[i].setAttr('lb', int(self.variable_lb[i]))
            x[i].setAttr('ub', int(self.variable_ub[i]))
        m.update()
        
        # set objective
        objective = quicksum(self.obj_coeff[i] * x[i] for i in range(self.n_variables))
        m.setObjective(objective, self.obj)        
        
        # add constraints
        if self.cons_df.shape[0]:
            ## add le constraints
            le_const = self.cons_df[self.cons_df['constraint type'] == '<=']
            le_const.reset_index(inplace=True, drop=True)
            for i in range(le_const.shape[0]):
                m.addConstr(quicksum(float(le_const.iloc[i,j])* x[j] for j in range(self.n_variables)) <= float(le_const['RHS values'][i]))
                
            ## add ge constraints
            ge_const = self.cons_df[self.cons_df['constraint type'] == '>=']
            ge_const.reset_index(inplace=True, drop=True)
            for i in range(ge_const.shape[0]):
                m.addConstr(quicksum(float(ge_const.iloc[i,j])* x[j] for j in range(self.n_variables)) >= float(ge_const['RHS values'][i]))
            ## add eq constraints
            eq_const = self.cons_df[self.cons_df['constraint type'] == '=']
            eq_const.reset_index(inplace=True, drop=True)
            for i in range(eq_const.shape[0]):
                m.addConstr(quicksum(float(eq_const.iloc[i,j])* x[j] for j in range(self.n_variables)) == float(eq_const['RHS values'][i]))
                

        def sensitivity_analysis():
            for v in m.getVars():
                print("For Variable " + v.VarName+ ":")
                print("Minimum value coefficient can take before the optimal decision changes "  + "is " + str(v.SAObjLow))
                print("Maximum value coefficient can take before the optimal decision changes "  + "is " + str(v.SAObjUp))
            
            for c in m.getConstrs():
                print("For constraint " + c.ConstrName+ ":")
                print("Shadow Price is " + str(c.pi))
                print("Minimum value RHS can take before the shadow price changes "  + "is " + str(c.SARHSLow))
                print("Maximum value RHS can take before the shadow price changes "  + "is " + str(c.SARHSUp))
                
        m.optimize()
        if run_sa==True:
            if "I" in self.variable_type or "B" in self.variable_type:
                raise NameError('There is at least one variable with type of integer/binary.')
            else:
                sensitivity_analysis()

        print(m.printAttr('X'))

    def ezprint(self, to='screen'):
        
        print("self")
        


if "__name__" == "__main__":
    m = EzGurobi()
    m.ezload("Model_Inputs.xlsx")
    m.ezrun()
# %%
