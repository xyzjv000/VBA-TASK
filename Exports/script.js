function totalValuesByTypeNew(data) {
    const margins = {};
    const names = [
      { type: "Capacity", name: "capacity" },
      { type: "Commission", name: "commission" },
      { type: "ESS", name: "ess" },
      { type: "LGC", name: "lgc" },
      { type: "Market Fees", name: "marketFees" },
      { type: "Network", name: "network" },
      { type: "Retail Margin", name: "retailMargin" },
      { type: "Revenue", name: "revenue" },
      { type: "STC", name: "stc" },
      { type: "Wholesale Energy", name: "wholesaleEnergy" },
    ];
  
    // Create a mapping of type names for easy lookup
    const typeMapping = Object.fromEntries(names.map(item => [item.name, item.type]));
  
    data.forEach((item) => {
      if (!margins[item.margin]) {
        // Initialize a new margin if not already present
        margins[item.margin] = {};
      }
  
      item.type.forEach((typeObj) => {
        // Use typeMapping to update the type value
        const updatedType = typeMapping[typeObj.name] || typeObj.name; // fallback to original name if not found
  
        if (margins[item.margin][updatedType]) {
          margins[item.margin][updatedType] += typeObj.value;
        } else {
          margins[item.margin][updatedType] = typeObj.value;
        }
      });
    });
  
    // Convert the result back to the required format
    return Object.keys(margins).map((margin) => {
      return {
        margin,
        type: Object.keys(margins[margin]).map((name) => ({
          name,
          value: margins[margin][name],
        })),
      };
    });
  }
  