// ============================================
// 🔻 Import Libraries Section
// ============================================
import * as React from "react";
// import fetch from "node-fetch";
import { IInputs } from "./generated/ManifestTypes";
import {
  DetailsList,
  IColumn,
  DetailsListLayoutMode,
  SearchBox,
  Selection,
  SelectionMode,
  PrimaryButton,
} from "@fluentui/react";

// ============================================
// 🔻 Type Declarations Section
// ============================================
// Declare 'DummyData' type (to be deleted)
interface DummyData {
  id: number;
  name: string;
}
// Declare 'RequestBody' type
interface RequestBody {
  account: string | undefined;
  contact: string[];
}

// ============================================
// 🔻 Component Input Interface Section (pass data and arguments using the 'context' keyword)
// ============================================
/* 
Video: https://youtu.be/R1hTz-T5feQ?si=JAccsVjHru1K8hZl.
Chat: https://chatgpt.com/c/680890cf-ab18-8010-ad6e-d31757052e66. Scroll to the way top for the info
*/
interface ListComponentControlProps {
  context: ComponentFramework.Context<IInputs>;
  notifyOutputChanged: () => void;
}

// ============================================
// 🔻 Functions Section (will soon be moved to a file)
// ============================================
// Fetch ScAccount Data
async function fetchScAccountsData(
  context: ComponentFramework.Context<IInputs>
) {
  try {
    const result = await context.webAPI.retrieveMultipleRecords(
      "crff8_scaccount",
      "?$select=crff8_scaccountnumber,crff8_scaccountname"
    );
    console.log("scaccount records:", result.entities);
    return result.entities;
  } catch (error) {
    console.log("NOOOOO");
    console.error("Error retrieving scaccount records:", error);
    return [];
  }
}
// Fetch ScContact Data where ScContact is NOT associated with ScAccount via association table
/* 
Originally, it was "Fetch ScContact Data where ScContact IS associated with ScAccount via association table", I changed operator from 'eq' to 'ne' but it didn't work,
it turned in to "Fetch ScContact Data where ScContact IS associated with OTHER ScAccount BUT NOT with the given ScAccount GUID via association table",
so I need to find another way which is this way below. 
*/
async function fetchStakeholdersDataAssociateNot(
  context: ComponentFramework.Context<IInputs>
) {
  try {
    const OpportunityGUID = context.parameters.sampleText.raw;
    console.log(`> Opportunity GUID Value: ${OpportunityGUID}`);
    const fetchXML = `<fetch>
            <entity name='crff8_stakeholder'>
              <attribute name='crff8_stakeholderid' />
              <attribute name='crff8_name' />
              <attribute name='crff8_contactinfo' />
              <link-entity name='crff8_stakeholder_opportunity'
                          from='crff8_stakeholderid'
                          to='crff8_stakeholderid'
                          link-type='outer'
                          alias='link'>
                <filter type='and'>
                  <condition attribute='opportunityid' operator='eq' value='${OpportunityGUID}' />
                </filter>
              </link-entity>
              <filter type='and'>
                <condition entityname='link' attribute='opportunityid' operator='null' />
              </filter>
            </entity>
          </fetch>`;
    const encodedFetchXML = encodeURIComponent(fetchXML);
    const result = await context.webAPI.retrieveMultipleRecords(
      "crff8_stakeholder",
      `?fetchXml=${encodedFetchXML}`
    );
    console.log("> Stakeholder associate not:", result.entities);
    return result.entities;
  } catch (error) {
    console.log(">> Error retrieving sccontact associate records:", error);
    console.error("Error retrieving sccontact associate records:", error);
    return [];
  }
}
// Create an HTTP fetch to the Power Automate to run the native 'Relate Row' action (not a Custom Action)
/* 
This Power Automate Flow is called 'flow-test-code-2' and it doesn't have an error case, will add it later or handle error case via code in here
*/
async function triggerRelateRowFlow(URL: string, body: RequestBody) {
  try {
    return [];
  } catch (error) {
    return [];
  }
}
// Sort the list by 'Name' from A-Z
const sortListByNameAZ = (a: unknown, b: unknown) => {
  const nameA = (a as { crff8_name?: string }).crff8_name?.toLowerCase() || "";
  const nameB = (b as { crff8_name?: string }).crff8_name?.toLowerCase() || "";
  return nameA.localeCompare(nameB);
};

// ============================================
// 🔻 Main Component Section (this is where the magic begin)
// ============================================
const ListComponentControl: React.FC<ListComponentControlProps> = ({
  context,
  notifyOutputChanged,
}) => {
  // Initialize dummy data with data (to be deleted)
  const dummyData: DummyData[] = [
    { id: 1, name: "John Doe" },
    { id: 2, name: "Jane Smith" },
    { id: 3, name: "Alice Johnson" },
    { id: 4, name: "Bob Brown" },
  ];

  // ---------------------------
  // State Variables (initialize "state" to hold and set/change value)
  // ---------------------------
  const [searchText, setSearchText] = React.useState<string>(""); // State to hold "searchText" and "setSearchText", no initial value
  const [scAccounts, setScAccounts] = React.useState<unknown[]>([]); // State to hold "scAccounts" and "setScAccounts", no initial value
  const [Stakeholders, setStakeholders] = React.useState<unknown[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [selection] = React.useState(
    new Selection({
      onSelectionChanged: () => {
        const selectedItems = selection.getSelection();
        if (selectedItems.length > 0) {
          console.log("Các dòng được chọn:");
          selectedItems.forEach((item, index) => {
            console.log(`Row ${index + 1}:`, item);
          });
        } else {
          console.log("Không có dòng nào được chọn.");
        }
      },
    })
  );

  // Run once when the page is render and run twice when context is updated (context is updated when the page is reloaded)
  React.useEffect(() => {
    const loadData = async () => {
      setIsLoading(true);
      const stakeholders = await fetchStakeholdersDataAssociateNot(context);
      setStakeholders(stakeholders);
      setIsLoading(false);
    };
    loadData();
    // const loadData = async () => {
    //   setIsLoading(true);
    //   const [account] = await Promise.all([
    //     fetchScAccountsData(context),
    //     fetchScContactsDataAssociateNot(context),
    //   ]);
    //   setScAccounts(account);
    //   setIsLoading(false);
    // };
    // loadData();
  }, [context]);

  // ---------------------------
  // Event Handlers
  // ---------------------------
  // Compute filtered items (with useMemo) whenever 'searchText' or 'scContacts' changes
  const filteredItems = React.useMemo(() => {
    const term = searchText.trim().toLowerCase(); // Get the search term from 'searchText'
    if (!term) {
      return (Stakeholders as unknown[]).slice().sort(sortListByNameAZ);
    }
    return (Stakeholders as unknown[])
      .filter((item) => {
        const stakeholder = item as {
          crff8_name?: string;
          crff8_stakeholderid?: string;
        };
        return (
          stakeholder.crff8_name?.toLowerCase().includes(term) || // Match on name or number
          stakeholder.crff8_stakeholderid?.toLowerCase().includes(term)
        );
      })
      .slice()
      .sort(sortListByNameAZ);
  }, [searchText, Stakeholders]); // Dependency array, if any of these variables change, it triggers this function, idk how to explain further
  // Create (many-many) associate between ScAccount and ScContact via button click
  const handleGetSelectedId = async () => {
    const selectedItems = selection.getSelection();
    // Throw an alert when the button is clicked with no selected row
    if (selectedItems.length === 0) {
      alert("Warning: No row is selected. Please select a row.");
      return;
    }
    // Extract the GUID ONLY from the contact object
    const stakeholderIds = selectedItems
      .map((item) => {
        const stakeholder = item as {
          crff8_name?: string;
          crff8_stakeholderid?: string;
        };
        return stakeholder.crff8_stakeholderid;
      })
      .filter((id): id is string => !!id);
    // Calling PA, careful because PA only has success cases, will create error handler in PA or in here
    console.log("> Opportunity GUID: ", context?.parameters?.sampleText.raw);
    console.log("> Selected stakeholders: ", stakeholderIds);
    setIsLoading(true);
    try {
      // Get the Env Var for the Associate Flow
      const envVarGuidResult = await context.webAPI.retrieveMultipleRecords(
        "environmentvariabledefinition",
        "?$filter=schemaname eq 'crff8_AssociateFlow'&$select=schemaname,environmentvariabledefinitionid"
      );
      if (envVarGuidResult.entities.length == 0) {
        throw new Error("Associate Flow Guid is empty!!!");
      }
      const envVarGuid =
        envVarGuidResult.entities[0]?.environmentvariabledefinitionid;
      const result = await context.webAPI.retrieveMultipleRecords(
        "environmentvariablevalue",
        `?$filter=_environmentvariabledefinitionid_value eq '${envVarGuid}'&$select=schemaname,value`
      );
      const envVarValue = result.entities[0].value;
      console.log("Env Var Associate Flow: ", envVarValue);

      const URL = envVarValue;
      // "https://prod-27.southeastasia.logic.azure.com:443/workflows/9a095ece2f71414ba3244cab5bd7d913/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=WvGhPBthO8aeiAFQYWY7d6QuI57U0AThSpSm0-eGgVY";
      const body = JSON.stringify({
        opportunity: context?.parameters?.sampleText.raw,
        stakeholder: stakeholderIds,
      });
      const response = await fetch(URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: body,
      });
      const data = await response.json();
      console.log("> Success", data);
      location.reload();
    } catch (error) {
      console.log(`❌ Failed to link contact`, error);
      console.error(`❌ Failed to link contact`, error);
    } finally {
      setIsLoading(false);
    }
  };

  // ---------------------------
  // Render Components
  // ---------------------------
  // Define columns used in the component
  const columns: IColumn[] = [
    {
      key: "column1",
      name: "Name", // Display name
      fieldName: "crff8_name", // This is where the data is mapped based on the column name/logical name
      minWidth: 50,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "column2",
      name: "Contact Info",
      fieldName: "crff8_contactinfo", // This is where the data is mapped based on the column name/logical name
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
    },
    {
      key: "column3",
      name: "GUID",
      fieldName: "crff8_stakeholderid", // This is where the data is mapped based on the column name/logical name
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
    },
  ];

  const title = context?.parameters?.sampleProperty.raw ?? "Unknown Title";
  const value = context?.parameters?.sampleText.raw ?? "Unknown Value";
  // The "loading overlay" component
  if (isLoading) {
    return (
      <div
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: "rgba(255, 255, 255, 0.7)",
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          alignItems: "center",
          zIndex: 9999,
        }}
      >
        <div
          style={{
            border: "6px solid #f3f3f3",
            borderTop: "6px solid #3498db",
            borderRadius: "50%",
            width: "50px",
            height: "50px",
            animation: "spin 1s linear infinite",
          }}
        />
        <p>Loading...</p>
        <style>{`
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
        `}</style>
      </div>
    );
  }
  // Main component
  return (
    <div style={{ padding: "16px" }}>
      {/* Header */}
      <div style={{ textAlign: "center", marginBottom: "20px" }}>
        <h3 style={{ margin: 0 }}>Opportunity Topic Name: {title}</h3>
        <p style={{ margin: 0 }}>Record ID: {value}</p>
      </div>
      {/* Search box */}
      <div style={{ width: "100%", maxWidth: "600px", margin: "0 auto 12px" }}>
        <SearchBox
          placeholder="Tìm theo Name hoặc GUID..."
          value={searchText} // This is for displaying UI only
          onChange={(_, newValue) => setSearchText(newValue || "")} // When the user change any values, it will update 'searchText'
          underlined={false}
        />
      </div>
      {/* Display list */}
      <div
        style={{
          width: "100%",
          margin: "0 auto",
          height: "300px",
          overflowY: "auto",
        }}
      >
        <DetailsList
          items={filteredItems} // Instead of dummy value, it will load based on filteredItems from the search
          columns={columns} // Used to columns we defined before
          setKey="filtered"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          selection={selection} // Enable selection of table trigger everytime there is a 'selection' state
          selectionPreservedOnEmptyClick={true}
          selectionMode={SelectionMode.multiple}
        />
      </div>
      {/* Divider*/}
      <div style={{ borderTop: "1px solid #ccc", margin: "12px 0" }}></div>
      {/* Row count section*/}
      <div
        style={{ textAlign: "left", marginTop: "12px", marginBottom: "12px" }}
      >
        {/* This line shows the row count */}
        <p style={{ margin: 0 }}>Rows: {filteredItems.length}</p>
      </div>
      {/* Button */}
      <div
        style={{ textAlign: "center", marginTop: "12px", marginBottom: "12px" }}
      >
        <PrimaryButton
          text="Get the GUID from selected row(s)"
          onClick={handleGetSelectedId}
        />
      </div>
    </div>
  );
};

// ============================================
// 🔻 Export Component
// ============================================
export default ListComponentControl;
