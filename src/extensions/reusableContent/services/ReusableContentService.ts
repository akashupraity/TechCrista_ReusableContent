// src/services/ReusableContentService.ts
// Import necessary modules from the PnP library to interact with SharePoint
import { sp } from "@pnp/sp";// PnPjs library for SharePoint interactions
import "@pnp/sp/webs";// Enables web-level operations
import "@pnp/sp/lists";// Enables list operations
import "@pnp/sp/items";// Enables item operations


// Define interfaces for the SharePoint item and content
// IContent defines the structure for reusable content in the application
export interface IContent {
  title: string;
  content: string;
  order: number;
  expand: boolean;
}
// ISharePointItem defines the structure of SharePoint items retrieved from the list
interface ISharePointItem {
  Title: string;
  Content: string;
  Order0: number;
  Expand: boolean;
}

// Create a service class to handle reusable content operations
// This class encapsulates the logic for fetching reusable content from SharePoint
export class ReusableContentService {
  // Method to fetch content items from the SharePoint "ReusableContent" list
  public static async getContentItems(): Promise<IContent[]> {
    try {
      // Fetch items from the "ReusableContent" list, selecting specific fields
      
      const items: ISharePointItem[] = await sp.web.lists
        .getByTitle("ReusableContent")// Access the specific list by its title
        .items
        .select("Title", "Content", "Order0", "Expand")// Specify fields to retrieve
        .orderBy("Order0", true) // Fetching in ascending order
        .get();// Execute the query to get the items

      // Map the SharePoint items to the application-defined IContent interface
      return items.map((item: ISharePointItem) => ({
        title: item.Title,
        content: item.Content,
        order: item.Order0,   // Ensure this is a number
        expand: item.Expand   // Capture the expand property
      }));

    } catch (error) {
        // Handle any errors that occur during the fetch operation
      console.error("Error fetching content items:", error);
      return []; // Return an empty array if an error occurs
    }
  }
}
