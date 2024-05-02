import numpy as np
import skfuzzy as fuzz
from skfuzzy import control as ctrl
import matplotlib.pyplot as plt
import pandas as pd

def plot_membership_functions(antecedent):
    plt.figure(figsize=(8, 6))
    x = antecedent.universe
    for key in antecedent.terms.keys():
        membership_degree = fuzz.interp_membership(x, antecedent[key].mf, x)
        plt.plot(x, membership_degree, label=key)
    plt.title(f'Membership Functions for {antecedent.label}')
    plt.xlabel(antecedent.label)
    plt.ylabel('Membership Degree')
    plt.legend()
    plt.grid(True)
    plt.show()




def fuzzyprogram(name, growth, pocfval, debtequityval, quickval, roassetsval, rocapitalval, buybackval):
    # Define input variables
    pocf = ctrl.Antecedent(np.arange(-20, 41, 1), 'price-to-operatingcf') # keep (negative means cant sustain itself) 15 range 
    debtequity = ctrl.Antecedent(np.arange(-1, 2.1, 0.1), 'debt-to-equity') # keep (is measure normal is 0-1 managable, higher than 1 bad)
    #how much debt the company consistents of (0.5 = half money is debt)
    quick = ctrl.Antecedent(np.arange(-1, 4.1, 0.1), 'quick-ratio') # keep higher is better (ability to pay back debt)
    roassets = ctrl.Antecedent(np.arange(-30, 51, 1), 'return-on-assets') # keep (50% and up is great)
    rocapital = ctrl.Antecedent(np.arange(-30, 50, 1), 'return-on-capital') # keep (50% and up is great)
    buyback = ctrl.Antecedent(np.arange(-30, 30, 1), 'buyback-yield') # keep (is a percentage, around 0 is average)

    # Define output variable
    bankruptcy = ctrl.Consequent(np.arange(0, 101, 1), 'bankruptcy')

    # Define the weights of the ratios
    pocfweight = 0.9
    debtequityweight = 1.3
    quickweight = 1.1
    roassetsweight = 1.2
    rocapialweight = 1.2
    buybackweight = 0.8

    # Define membership functions for input variables (Gaussian)
    pocf['Very Low'] = fuzz.gaussmf(pocf.universe, -20, 6.4*pocfweight)
    pocf['Low'] = fuzz.gaussmf(pocf.universe, -5, 6.4*pocfweight)
    pocf['Average'] = fuzz.gaussmf(pocf.universe, 10, 6.4*pocfweight)
    pocf['High'] = fuzz.gaussmf(pocf.universe, 25, 6.4*pocfweight)
    pocf['Very High'] = fuzz.gaussmf(pocf.universe, 40, 6.4*pocfweight)

    debtequity['Very Low'] = fuzz.gaussmf(debtequity.universe, -1, 0.32*debtequityweight)
    debtequity['Low'] = fuzz.gaussmf(debtequity.universe, -0.25, 0.32*debtequityweight)
    debtequity['Average'] = fuzz.gaussmf(debtequity.universe, 0.5, 0.32*debtequityweight)
    debtequity['High'] = fuzz.gaussmf(debtequity.universe, 1.25, 0.32*debtequityweight)
    debtequity['Very High'] = fuzz.gaussmf(debtequity.universe, 2, 0.32*debtequityweight)

    quick['Very Low'] = fuzz.gaussmf(quick.universe, -1, 0.55*quickweight)
    quick['Low'] = fuzz.gaussmf(quick.universe, 0.25, 0.55*quickweight)
    quick['Average'] = fuzz.gaussmf(quick.universe, 1.5, 0.55*quickweight)
    quick['High'] = fuzz.gaussmf(quick.universe, 2.75, 0.55*quickweight)
    quick['Very High'] = fuzz.gaussmf(quick.universe, 4, 0.55*quickweight)

    roassets['Very Low'] = fuzz.gaussmf(roassets.universe, -30, 8.5*roassetsweight)
    roassets['Low'] = fuzz.gaussmf(roassets.universe, -10, 8.5*roassetsweight)
    roassets['Average'] = fuzz.gaussmf(roassets.universe, 10, 8.5*roassetsweight)
    roassets['High'] = fuzz.gaussmf(roassets.universe, 30, 8.5*roassetsweight)
    roassets['Very High'] = fuzz.gaussmf(roassets.universe, 50, 8.5*roassetsweight)

    rocapital['Very Low'] = fuzz.gaussmf(rocapital.universe, -30, 8.5*rocapialweight)
    rocapital['Low'] = fuzz.gaussmf(rocapital.universe, -10, 8.5*rocapialweight)
    rocapital['Average'] = fuzz.gaussmf(rocapital.universe, 10, 8.5*rocapialweight)
    rocapital['High'] = fuzz.gaussmf(rocapital.universe, 30, 8.5*rocapialweight)
    rocapital['Very High'] = fuzz.gaussmf(rocapital.universe, 50, 8.5*rocapialweight)

    buyback['Very Low'] = fuzz.gaussmf(buyback.universe, -30, 5*buybackweight)
    buyback['Low'] = fuzz.gaussmf(buyback.universe, -15, 5*buybackweight)
    buyback['Average'] = fuzz.gaussmf(buyback.universe, 0, 5*buybackweight)
    buyback['High'] = fuzz.gaussmf(buyback.universe, 15, 5*buybackweight)
    buyback['Very High'] = fuzz.gaussmf(buyback.universe, 30, 5*buybackweight)

    # Define the membership for the output variable
    bankruptcy['Very Low'] = fuzz.gaussmf(bankruptcy.universe, 0, 10.5)
    bankruptcy['Low'] = fuzz.gaussmf(bankruptcy.universe, 25, 10.5)
    bankruptcy['Average'] = fuzz.gaussmf(bankruptcy.universe, 50, 10.5)
    bankruptcy['High'] = fuzz.gaussmf(bankruptcy.universe, 75, 10.5)
    bankruptcy['Very High'] = fuzz.gaussmf(bankruptcy.universe, 100, 10.5)
    
    # Define fuzzy rules

    rulepocf1 = ctrl.Rule(pocf['Very Low'], bankruptcy['Very High'])
    rulepocf2 = ctrl.Rule(pocf['Low'], bankruptcy['Very High'])
    rulepocf3 = ctrl.Rule(pocf['Average'], bankruptcy['Average'])
    rulepocf4 = ctrl.Rule(pocf['High'], bankruptcy['Low'])
    rulepocf5 = ctrl.Rule(pocf['Very High'], bankruptcy['Very Low'])

    listpocf = [rulepocf1, rulepocf2, rulepocf3, rulepocf4, rulepocf5]

    ruledebtequity1 = ctrl.Rule(debtequity['Very Low'], bankruptcy['Very High'])
    ruledebtequity2 = ctrl.Rule(debtequity['Low'], bankruptcy['High'])
    ruledebtequity3 = ctrl.Rule(debtequity['Average'], bankruptcy['Average'])
    ruledebtequity4 = ctrl.Rule(debtequity['High'], bankruptcy['Very High'])
    ruledebtequity5 = ctrl.Rule(debtequity['Very High'], bankruptcy['High'])

    listdebtequity = [ruledebtequity1, ruledebtequity2, ruledebtequity3, ruledebtequity4, ruledebtequity5]

    rulequick1 = ctrl.Rule(quick['Very Low'], bankruptcy['Very High'])
    rulequick2 = ctrl.Rule(quick['Low'], bankruptcy['High'])
    rulequick3 = ctrl.Rule(quick['Average'], bankruptcy['Average'])
    rulequick4 = ctrl.Rule(quick['High'], bankruptcy['Very Low'])
    rulequick5 = ctrl.Rule(quick['Very High'], bankruptcy['Very Low'])

    listquick = [rulequick1, rulequick2, rulequick3, rulequick4, rulequick5]

    ruleroassets1 = ctrl.Rule(roassets['Very Low'], bankruptcy['Very High'])
    ruleroassets2 = ctrl.Rule(roassets['Low'], bankruptcy['High'])
    ruleroassets3 = ctrl.Rule(roassets['Average'], bankruptcy['Average'])
    ruleroassets4 = ctrl.Rule(roassets['High'], bankruptcy['Low'])
    ruleroassets5 = ctrl.Rule(roassets['Very High'], bankruptcy['Very Low'])

    listroassets = [ruleroassets1, ruleroassets2, ruleroassets3, ruleroassets4, ruleroassets5]

    rulerocapial1 = ctrl.Rule(rocapital['Very Low'], bankruptcy['Very High'])
    rulerocapial2 = ctrl.Rule(rocapital['Low'], bankruptcy['High'])
    rulerocapial3 = ctrl.Rule(rocapital['Average'], bankruptcy['Average'])
    rulerocapial4 = ctrl.Rule(rocapital['High'], bankruptcy['Low'])
    rulerocapial5 = ctrl.Rule(rocapital['Very High'], bankruptcy['Very Low'])

    listrocapial = [rulerocapial1, rulerocapial2, rulerocapial3, rulerocapial4, rulerocapial5]

    rulebuyback1 = ctrl.Rule(buyback['Very Low'], bankruptcy['Very High'])
    rulebuyback2 = ctrl.Rule(buyback['Low'], bankruptcy['High'])
    rulebuyback3 = ctrl.Rule(buyback['Average'], bankruptcy['Average'])
    rulebuyback4 = ctrl.Rule(buyback['High'], bankruptcy['Low'])
    rulebuyback5 = ctrl.Rule(buyback['Very High'], bankruptcy['Low'])

    listbuyback = [rulebuyback1, rulebuyback2, rulebuyback3, rulebuyback4, rulebuyback5]
    
    rule4 = ctrl.Rule(pocf['Low'] & quick['Low'], bankruptcy['High'])
    rule2 = ctrl.Rule(pocf['Average'] & debtequity['Average'], bankruptcy['Average'])
    rule3 = ctrl.Rule(pocf['High'] | debtequity['High'], bankruptcy['High'])

    rule12 = ctrl.Rule(pocf['Low'] & debtequity['Very High'], bankruptcy['Very High'])
    rule1 = ctrl.Rule(pocf['Low'] & debtequity['High'], bankruptcy['Very High'])
    rule8 = ctrl.Rule(pocf['Low'] & debtequity['Low'], bankruptcy['Very High'])
    rule10 = ctrl.Rule(pocf['Low'] & debtequity['Very Low'], bankruptcy['Very High'])
    rule9 = ctrl.Rule(pocf['Very Low'] & debtequity['Low'], bankruptcy['Very High'])
    rule11 = ctrl.Rule(pocf['Very Low'] & debtequity['Very Low'], bankruptcy['Very High'])
    rule7 = ctrl.Rule(pocf['Very Low'] & debtequity['High'], bankruptcy['Very High'])
    rule7 = ctrl.Rule(pocf['Very Low'] & debtequity['Very High'], bankruptcy['Very High'])
    rule5 = ctrl.Rule(pocf['High'] & debtequity['Average'], bankruptcy['Low'])
    rule6 = ctrl.Rule(pocf['Very High'] & debtequity['Average'], bankruptcy['Very Low'])
    
    rule13 = ctrl.Rule(quick['Very High'] & debtequity['Average'], bankruptcy['Very Low'])
    rule14 = ctrl.Rule(quick['High'] & debtequity['Average'], bankruptcy['Very Low'])
    rule15 = ctrl.Rule(quick['Average'] & debtequity['Average'], bankruptcy['Low'])
    rule16 = ctrl.Rule(quick['Low'] & debtequity['Average'], bankruptcy['High'])
    rule17 = ctrl.Rule(quick['Very Low'] & debtequity['Average'], bankruptcy['High'])
    rule18 = ctrl.Rule(quick['Low'] & debtequity['High'], bankruptcy['Very High'])
    rule19 = ctrl.Rule(quick['Very Low'] & debtequity['High'], bankruptcy['Very High'])
    rule20 = ctrl.Rule(quick['Low'] & debtequity['Very High'], bankruptcy['Very High'])
    rule21 = ctrl.Rule(quick['Very Low'] & debtequity['Very High'], bankruptcy['Very High'])
    rule22 = ctrl.Rule(quick['Very High'] & debtequity['High'], bankruptcy['Average'])
    rule23 = ctrl.Rule(quick['High'] & debtequity['High'], bankruptcy['High'])
    rule24 = ctrl.Rule(quick['High'] & debtequity['Very High'], bankruptcy['High'])
    rule25 = ctrl.Rule(quick['Very High'] & debtequity['Very High'], bankruptcy['High'])
    
    listcombined = [rule1, rule2, rule3, rule4, rule5, rule6, rule7, rule8, rule9, rule10, rule11, rule12,
                   rule13, rule14, rule15, rule16, rule17, rule18, rule19, rule20, rule21, rule22, rule23,
                   rule24, rule25]
    
    rule26 = ctrl.Rule(roassets['Very High'] & debtequity['Very High'], bankruptcy['High'])
    rule27 = ctrl.Rule(roassets['Very High'] & debtequity['High'], bankruptcy['Average'])
    rule28 = ctrl.Rule(roassets['Very High'] & debtequity['Average'], bankruptcy['Very Low'])
    rule29 = ctrl.Rule(roassets['Very High'] & debtequity['Low'], bankruptcy['Average'])
    rule30 = ctrl.Rule(roassets['Very High'] & debtequity['Very Low'], bankruptcy['High'])
    rule31 = ctrl.Rule(roassets['High'] & debtequity['Very High'], bankruptcy['High'])
    rule32 = ctrl.Rule(roassets['High'] & debtequity['High'], bankruptcy['High'])
    rule33 = ctrl.Rule(roassets['High'] & debtequity['Average'], bankruptcy['Very Low'])
    rule34 = ctrl.Rule(roassets['High'] & debtequity['Low'], bankruptcy['High'])
    rule35 = ctrl.Rule(roassets['High'] & debtequity['Very Low'], bankruptcy['Very High'])
    rule36 = ctrl.Rule(roassets['Average'] & debtequity['Very High'], bankruptcy['Very High'])
    rule37 = ctrl.Rule(roassets['Average'] & debtequity['High'], bankruptcy['High'])
    rule38 = ctrl.Rule(roassets['Average'] & debtequity['Average'], bankruptcy['Average'])
    rule39 = ctrl.Rule(roassets['Average'] & debtequity['Low'], bankruptcy['High'])
    rule40 = ctrl.Rule(roassets['Average'] & debtequity['Very Low'], bankruptcy['Very High'])
    rule41 = ctrl.Rule(roassets['Low'] & debtequity['Very High'], bankruptcy['Very High'])
    rule42 = ctrl.Rule(roassets['Low'] & debtequity['High'], bankruptcy['Very High'])
    rule43 = ctrl.Rule(roassets['Low'] & debtequity['Average'], bankruptcy['Average'])
    rule44 = ctrl.Rule(roassets['Low'] & debtequity['Low'], bankruptcy['High'])
    rule45 = ctrl.Rule(roassets['Low'] & debtequity['Very Low'], bankruptcy['Very High'])
    rule46 = ctrl.Rule(roassets['Very Low'] & debtequity['Very High'], bankruptcy['Very High'])
    rule47 = ctrl.Rule(roassets['Very Low'] & debtequity['High'], bankruptcy['Very High'])
    rule48 = ctrl.Rule(roassets['Very Low'] & debtequity['Average'], bankruptcy['High'])
    rule49 = ctrl.Rule(roassets['Very Low'] & debtequity['Low'], bankruptcy['High'])
    rule50 = ctrl.Rule(roassets['Very Low'] & debtequity['Very Low'], bankruptcy['Very High'])
    
    listcombined1 = [rule26, rule27, rule28, rule29, rule30, rule31, rule32, rule33, rule34, rule35, rule36, rule37,
                    rule38, rule39, rule40, rule41, rule42, rule43, rule44, rule45, rule46, rule47, rule48, rule49,
                    rule50]
    
    rule51 = ctrl.Rule(rocapital['Very High'] & debtequity['Very High'], bankruptcy['High'])
    rule52 = ctrl.Rule(rocapital['Very High'] & debtequity['High'], bankruptcy['Average'])
    rule53 = ctrl.Rule(rocapital['Very High'] & debtequity['Average'], bankruptcy['Very Low'])
    rule54 = ctrl.Rule(rocapital['Very High'] & debtequity['Low'], bankruptcy['Average'])
    rule55 = ctrl.Rule(rocapital['Very High'] & debtequity['Very Low'], bankruptcy['High'])
    rule56 = ctrl.Rule(rocapital['High'] & debtequity['Very High'], bankruptcy['High'])
    rule57 = ctrl.Rule(rocapital['High'] & debtequity['High'], bankruptcy['High'])
    rule58 = ctrl.Rule(rocapital['High'] & debtequity['Average'], bankruptcy['Very Low'])
    rule59 = ctrl.Rule(rocapital['High'] & debtequity['Low'], bankruptcy['High'])
    rule60 = ctrl.Rule(rocapital['High'] & debtequity['Very Low'], bankruptcy['Very High'])
    rule61 = ctrl.Rule(rocapital['Average'] & debtequity['Very High'], bankruptcy['Very High'])
    rule62 = ctrl.Rule(rocapital['Average'] & debtequity['High'], bankruptcy['High'])
    rule63 = ctrl.Rule(rocapital['Average'] & debtequity['Average'], bankruptcy['Average'])
    rule64 = ctrl.Rule(rocapital['Average'] & debtequity['Low'], bankruptcy['High'])
    rule65 = ctrl.Rule(rocapital['Average'] & debtequity['Very Low'], bankruptcy['Very High'])
    rule66 = ctrl.Rule(rocapital['Low'] & debtequity['Very High'], bankruptcy['Very High'])
    rule67 = ctrl.Rule(rocapital['Low'] & debtequity['High'], bankruptcy['Very High'])
    rule68 = ctrl.Rule(rocapital['Low'] & debtequity['Average'], bankruptcy['Average'])
    rule69 = ctrl.Rule(rocapital['Low'] & debtequity['Low'], bankruptcy['High'])
    rule70 = ctrl.Rule(rocapital['Low'] & debtequity['Very Low'], bankruptcy['Very High'])
    rule71 = ctrl.Rule(rocapital['Very Low'] & debtequity['Very High'], bankruptcy['Very High'])
    rule72 = ctrl.Rule(rocapital['Very Low'] & debtequity['High'], bankruptcy['Very High'])
    rule73 = ctrl.Rule(rocapital['Very Low'] & debtequity['Average'], bankruptcy['High'])
    rule74 = ctrl.Rule(rocapital['Very Low'] & debtequity['Low'], bankruptcy['High'])
    rule75 = ctrl.Rule(rocapital['Very Low'] & debtequity['Very Low'], bankruptcy['Very High'])
    
    listcombined2 = [rule51, rule52, rule53, rule54, rule55, rule56, rule57, rule58, rule59, rule60, rule61, rule62,
                    rule63, rule64, rule65, rule66, rule67, rule68, rule69, rule70, rule71, rule72, rule73, rule74,
                    rule75]
    
    allcombined = listcombined + listcombined1 + listcombined2

    allrules = listpocf + listdebtequity + listquick + listroassets + listrocapial + listbuyback + allcombined

    # Create control system
    bankruptcy_ctrl = ctrl.ControlSystem(allrules)

    # Define simulation
    bankruptcy_sim = ctrl.ControlSystemSimulation(bankruptcy_ctrl)

    # Pass inputs to the simulation
    bankruptcy_sim.input['price-to-operatingcf'] = pocfval
    bankruptcy_sim.input['debt-to-equity'] = debtequityval
    bankruptcy_sim.input['quick-ratio'] = quickval
    bankruptcy_sim.input['return-on-assets'] = roassetsval
    bankruptcy_sim.input['return-on-capital'] = rocapitalval
    bankruptcy_sim.input['buyback-yield'] = buybackval


    # Compute the result
    bankruptcy_sim.compute()

    # Print the result
    print(name, "Growth: ", growth, "\nChance of bankruptcy:", str(bankruptcy_sim.output['bankruptcy']))
    print("\n")


    # Plots membership functions for variable
    # plot_membership_functions(pocf)
    # plot_membership_functions(debtequity)
    # plot_membership_functions(quick)
    # plot_membership_functions(roassets)
    # plot_membership_functions(rocapital)
    # plot_membership_functions(buyback)
    # plot_membership_functions(bankruptcy)



def read_spreadsheet(filename, category_names):
    """
    Read values from a spreadsheet file where each column represents a category and each row contains values.
    
    Parameters:
        filename (str): The filename of the spreadsheet file.
        category_names (list): A list of category names.
        
    Returns:
        tuple: A tuple containing a list of category names and a list of rows, where each row includes the name column.
    """
    # Read the spreadsheet file
    data = pd.read_excel(filename)
    
    # Select only the columns with the specified category names
    data = data[[data.columns[0]] + category_names]
    
    # Prepare rows with names
    rows = []
    for idx, row in data.iterrows():
        values = list(row)
        # Convert the final 6 values to floats
        for i in range(len(values) - 6, len(values)):
            value = values[i]
            # Check if the value is not already a float
            if not isinstance(value, float):
                # Remove "%" symbols
                value = value.replace("%", "")
                # Check if the value is "-"
                if value == "-":
                    value = "0"  # Replace "-" with "0"
                # Convert to float
                value = float(value)
                
                values[i] = value
        rows.append(values)

    # Extract category names including the name column
    category_names_with_name = ["Name"] + category_names
    
    return category_names_with_name, rows


#filename = "technologyCompiled.xlsx"
filename = "bankruptciesCompiled.xlsx"
#filename = "Bankrupcies/Bird Global_Ratios.xlsx"
category_names = ["Market Cap Growth", "P/OCF Ratio", "Debt / Equity Ratio", "Quick Ratio", "Return on Assets (ROA)", "Return on Capital (ROIC)", "Buyback Yield / Dilution"]
category_names_with_name, rows = read_spreadsheet(filename, category_names)

for row in rows:
    fuzzyprogram(*row)
    #print(*row)
    
#fuzzyprogram("Liberty Broadband Corporation", "-11%", -12.86, 0.16, 1.09, -47.70, -66.91, -125.5)

#Very Low and Low: x = 12.5
#Low and Average: x = 37.5
#Average and High: x = 62.5
#High and Very High: x = 87.5