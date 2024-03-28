import numpy as np
import skfuzzy as fuzz
from skfuzzy import control as ctrl
import matplotlib.pyplot as plt

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

# Define input variables
PE_ratio = ctrl.Antecedent(np.arange(0, 50, 1), 'temperature')
humidity = ctrl.Antecedent(np.arange(0, 101, 1), 'humidity')

# Define output variable
fan_speed = ctrl.Consequent(np.arange(0, 101, 1), 'fan_speed')

# Define membership functions for input variables (Gaussian)

PE_ratio['Very Low'] = fuzz.gaussmf(PE_ratio.universe, 0, 10)
PE_ratio['Low'] = fuzz.gaussmf(PE_ratio.universe, 15, 10)
PE_ratio['Medium'] = fuzz.gaussmf(PE_ratio.universe, 27.5, 10)
PE_ratio['High'] = fuzz.gaussmf(PE_ratio.universe, 40, 10)
PE_ratio['Very High'] = fuzz.gaussmf(PE_ratio.universe, 50, 10)

humidity['low'] = fuzz.gaussmf(humidity.universe, 0, 20)
humidity['medium'] = fuzz.gaussmf(humidity.universe, 50, 20)
humidity['high'] = fuzz.gaussmf(humidity.universe, 100, 20)

# Define membership functions for output variable (Gaussian)
fan_speed['low'] = fuzz.gaussmf(fan_speed.universe, 0, 20)
fan_speed['medium'] = fuzz.gaussmf(fan_speed.universe, 50, 20)
fan_speed['high'] = fuzz.gaussmf(fan_speed.universe, 100, 20)

# Plot membership functions for input variables
plot_membership_functions(PE_ratio)
plot_membership_functions(humidity)
plot_membership_functions(fan_speed)  # For output variable



# Define fuzzy rules
rule1 = ctrl.Rule(PE_ratio['low'] | humidity['low'], fan_speed['low'])
rule2 = ctrl.Rule(PE_ratio['medium'] & humidity['medium'], fan_speed['medium'])
rule3 = ctrl.Rule(PE_ratio['high'] | humidity['high'], fan_speed['high'])


# Create control system
fan_ctrl = ctrl.ControlSystem([rule1, rule2, rule3])

# Define simulation
fan_sim = ctrl.ControlSystemSimulation(fan_ctrl)

# Pass inputs to the simulation
fan_sim.input['PE_ratio'] = 30
fan_sim.input['humidity'] = 90

# Compute the result
fan_sim.compute()

# Print the result
print(fan_sim.output['fan_speed'])
