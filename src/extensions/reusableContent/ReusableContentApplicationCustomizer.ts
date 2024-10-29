import { override } from '@microsoft/decorators';// Importing decorators to allow method overriding
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';// Base class for SPFx application customizers
import { sp } from "@pnp/sp";// Importing the PnPjs library for SharePoint operations
import { ReusableContentService, IContent } from './services/ReusableContentService'; // Importing the service for reusable content operations


// Import the SCSS module
import styles from './IReusableContentApplicationCustomizer.module.scss';

// Define properties for the application customizer (currently empty)
export interface IReusableContentApplicationCustomizerProperties {}

// Create a class that extends BaseApplicationCustomizer to implement custom behavior
export default class ReusableContentApplicationCustomizer
  extends BaseApplicationCustomizer<IReusableContentApplicationCustomizerProperties> {

    // Array to store the fetched content items
  private _contentItems: IContent[] = [];

  // Override the onInit method to perform initialization tasks
  @override
  public async onInit(): Promise<void> {

    // Set up the PnP context for SharePoint operations
    sp.setup({ spfxContext: this.context as any });

    // Fetch content items using the service
    this._contentItems = await ReusableContentService.getContentItems();
    // Render the icon for accessing reusable content
    this._renderIcon();
  }

  // Method to render an icon on the page
  private _renderIcon(): void {
    const iconContainer = document.createElement("div");
    // Create a button to serve as the icon for opening the dialog
    iconContainer.innerHTML = `
      <button id="reusableContentIcon" class="${styles.iconButton}">
        📄
      </button>
    `;
    // Append the icon container to the body of the document
    document.body.appendChild(iconContainer);

    // Add a click event listener to the icon button to toggle the dialog
    const iconButton = document.getElementById("reusableContentIcon");
    if (iconButton) {
      iconButton.addEventListener("click", () => this._toggleDialog());
    }
  }

  // Method to toggle the display of the dialog containing reusable content
  private _toggleDialog(): void {
    const dialogContainer = document.createElement("div");
    dialogContainer.id = "reusableContentDialog";
    dialogContainer.className = styles.dialogContainer;

    const contentWrapper = document.createElement("div");
    // Populate the dialog with content
    contentWrapper.className = styles.contentWrapper;
    contentWrapper.innerHTML = this._getDialogContent();

    dialogContainer.appendChild(contentWrapper);

    // Create a close button for the dialog
    const closeButton = document.createElement("button");
    closeButton.innerText = "Close";
    closeButton.className = styles.closeButton;
    // Remove the dialog from the DOM when the close button is clicked
    closeButton.onclick = () => {
      document.body.removeChild(dialogContainer);
    };

    // Append the close button and the dialog container to the document body
    dialogContainer.appendChild(closeButton);
    document.body.appendChild(dialogContainer);

    // Set up accordion functionality for the dialog content
    this._setupAccordionToggle();
  }

  // Method to set up accordion functionality for displaying content items
  private _setupAccordionToggle(): void {
    this._contentItems.forEach(item => {
      // Get the accordion header element by ID
      const accordionHeader = document.getElementById(`${item.title.replace(/\s+/g, '-')}-header`);
      if (accordionHeader) {
        // Add a click event listener to toggle the accordion
        accordionHeader.addEventListener('click', () => this.toggleAccordion(item.title));

        // Automatically expand items based on the Expand property
        if (item.expand) {
          const contentDiv = document.getElementById(`${item.title.replace(/\s+/g, '-')}-content`);
          if (contentDiv) {
            contentDiv.style.display = "block"; // Show content if Expand is true
          }
        }
      }
    });
  }

  // Method to generate the HTML content for the dialog
  private _getDialogContent(): string {
    let contentHtml = `<h2 class="${styles.contentHeader}">Reusable Content</h2>`;
    
    // Loop through the fetched content items and build the HTML structure
    this._contentItems.forEach(item => {
      contentHtml += `
        <div style="border-bottom: 1px solid #ddd; padding: 10px 0;">
          <div id="${item.title.replace(/\s+/g, '-')}-header" class="${styles.accordionHeader}">
            <h3 style="margin: 0; display: flex; align-items: center; width:100%; justify-content: space-between">
              ${item.title}
              <span class="${styles.copyContentIcon}" data-content="${item.content}">
                📋
              </span>
            </h3>
          </div>
          <div id="${item.title.replace(/\s+/g, '-')}-content" class="${styles.accordionContent}">
            ${item.content}
          </div>
        </div>
      `;
    });
  
    return contentHtml;// Return the generated HTML
  }
  
  // Method to toggle the display of accordion content
  private toggleAccordion(title: string): void {
    const contentDiv = document.getElementById(`${title.replace(/\s+/g, '-')}-content`);
    if (contentDiv) {
      const isVisible = contentDiv.style.display === "block";
      contentDiv.style.display = isVisible ? "none" : "block";

      // Copy icon functionality
      if (!isVisible) {
        const copyIcon = contentDiv.previousElementSibling.querySelector(`.${styles.copyContentIcon}`);
        if (copyIcon) {
          // Add click event to copy content to clipboard
          copyIcon.addEventListener('click', () => {
            const content = copyIcon.getAttribute('data-content');
            if (content) {
              this.copyToClipboard(content);
            }
          });
        }
      }
    }
  }
// Method to copy content to clipboard
  private copyToClipboard(content: string): void {
    navigator.clipboard.writeText(content).then(() => {
      alert("Content copied to clipboard!");// Notify user of successful copy
    }).catch(err => {
      console.error("Failed to copy content: ", err);// Log error if copy fails
    });
  }
}
