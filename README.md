# CATIA V5 Automation Cheat Sheet (C#, Windows Forms)

## Introduction

This README serves as a quick reference guide for automating CATIA V5 using C# and Windows Forms. It covers essential tasks such as setting up CATIA automation, creating and modifying geometric elements, working with assemblies, error handling, and closing the CATIA application.

## Setting Up CATIA V5 Automation:

- **Add Reference to CATIA Type Library:**
```csharp
  using INFITF;
  using MECMOD;
  using ProductStructureTypeLib;
  using HybridShapeTypeLib;
```

- **Initialize CATIA Application:**
```csharp
  INFITF.Application catiaApp = null;
  catiaApp = (INFITF.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("CATIA.Application");
```

- **Create New Part Document:**
```csharp
  PartDocument partDoc = (PartDocument)catiaApp.Documents.Add("Part");
```

- **Accessing CATIA Part and Features:**
```csharp
  Part part = partDoc.Part;
  HybridBodies hybridBodies = part.HybridBodies;
```

## Creating Geometric Elements:

- **Create Point:**
```csharp
  HybridShapeFactory hybridShapeFactory = (HybridShapeFactory)part.HybridShapeFactory;
  Point point = hybridShapeFactory.AddNewPointCoord(x, y, z);
  part.Update();
```

- **Create Line:**
```csharp
  HybridShapeLinePtPt line = hybridShapeFactory.AddNewLinePtPt(point1, point2);
  part.Update();
```

- **Create Circle:**
```csharp
  HybridShapeCircle circle = hybridShapeFactory.AddNewCircleCtrRad(pointCenter, normal, radius);
  part.Update();
```

## Modifying Geometric Elements:

- **Translate/Rotate/Scale:**
```csharp
  HybridShapeTranslation translation = hybridShapeFactory.AddNewTranslation(line, translationVector);
  translation.Value = translationValue;
  part.Update();
```

- **Edit Parameters:**
```csharp
  Parameters parameters = part.Parameters;
  Parameter param = parameters.Item("ParameterName");
  param.Value = newValue;
  part.Update();
```

## Working with Assemblies:

- **Create Assembly Document:**
```csharp
  ProductDocument productDoc = (ProductDocument)catiaApp.Documents.Add("Product");
```

- **Insert Part into Assembly:**
```csharp
  Product product = productDoc.Product;
  product.AddItem(partDoc);
```

- **Manipulate Assembly Components:**
```csharp
  Products products = product.Products;
  Product component = products.Item(index);
  // Manipulate component
```

## Error Handling:

```csharp
  try
  {
      // CATIA operations
  }
  catch (Exception ex)
  {
      MessageBox.Show("Error: " + ex.Message);
  }
```

## Closing CATIA Application:
```csharp
  if (catiaApp != null)
  {
      catiaApp.Quit();
      catiaApp = null;
  }
```

##Example:
```csharp
// Below is an example code snippet demonstrating how to accomplish this:

using System;
using System.Windows.Forms;
using INFITF;
using MECMOD;
using ProductStructureTypeLib;

namespace CatiaAutomation
{
    public partial class MainForm : Form
    {
        private INFITF.Application catiaApp;
        private ProductDocument productDocument;
        private PartDocument partDocument;

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnCreatePart_Click(object sender, EventArgs e)
        {
            // Connect to CATIA
            try
            {
                catiaApp = (INFITF.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("CATIA.Application");
            }
            catch
            {
                MessageBox.Show("CATIA is not running or accessible.");
                return;
            }

            // Create a new part document
            catiaApp.Documents.Add("Part");

            // Activate the product document
            productDocument = (ProductDocument)catiaApp.ActiveDocument;

            // Get the root product
            Product rootProduct = productDocument.Product;

            // Create a new part
            partDocument = (PartDocument)catiaApp.ActiveDocument;
            Part part = partDocument.Part;

            // Add the part to the product structure
            ProductStructureFactory productStructureFactory = (ProductStructureFactory)productDocument.GetItem("ProductStructureFactory");
            productStructureFactory.AddComponent(part, rootProduct);
            
            // Refresh the product structure
            productDocument.Update();
            
            MessageBox.Show("New part created and added to the product.");
        }
    }
}
```
