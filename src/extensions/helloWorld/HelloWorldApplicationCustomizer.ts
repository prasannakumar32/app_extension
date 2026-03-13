import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import * as strings from 'HelloWorldApplicationCustomizerStrings';
import styles from './AppCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const LOG_SOURCE: string = 'HelloWorldApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldApplicationCustomizerProperties {
  Top: string;
  Bottom: string;
}

interface ISearchResultRow {
  Cells: { Key: string; Value: string; ValueType: string }[];
}

interface ISearchResponse {
  PrimaryQueryResult: {
    RelevantResults: {
      Table: {
        Rows: ISearchResultRow[];
      };
    };
  };
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HelloWorldApplicationCustomizer
  extends BaseApplicationCustomizer<IHelloWorldApplicationCustomizerProperties> {

  // These have been added
  private _topPlaceholder: PlaceholderContent | undefined;
  private _handleClickOutside?: (event: MouseEvent) => void;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // Wait for the placeholders to be created (or handle them being changed) and then
    // render.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);

    // Initial render call
    this._renderPlaceHolders();

    return Promise.resolve();
  }

  private _renderPlaceHolders(): void {
    // CRITICAL: Global check to prevent duplicate search bars across multiple extension instances
    if (document.getElementById('customSearchInput')) {
      return;
    }

    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this.onDispose.bind(this) }
      );
    }

    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      this._topPlaceholder.domElement.innerHTML = `
      <div class="${styles.app}">
        <div class="${styles.top}">
          <div class="${styles.searchContainer}">
            <div class="${styles.searchBox}">
              <input type="text" id="customSearchInput" class="${styles.searchInput}" placeholder="Search this site...">
              <button id="customClearButton" class="${styles.clearButton}" style="display:none;" title="Clear Search">
                <i class="ms-Icon ms-Icon--Clear" aria-hidden="true"></i>
              </button>
              <button id="customSearchButton" class="${styles.searchButton}">Search</button>
            </div>
            <div id="customSearchResults" class="${styles.resultsContainer}">
                <div class="${styles.resultsHeader}">
                    <span class="${styles.resultsTitle}">Search Results</span>
                    <button id="closeSearchResults" class="${styles.closeButton}" title="Close">
                        <i class="ms-Icon ms-Icon--Cancel" aria-hidden="true"></i>
                    </button>
                </div>
                <div id="resultsList" class="${styles.resultsList}"></div>
            </div>
          </div>
        </div>
      </div>`;

      this._attachSearchEvent();
    }
  }

  private _attachSearchEvent(): void {
    const searchButton = document.getElementById('customSearchButton');
    const searchInput = document.getElementById('customSearchInput') as HTMLInputElement;
    const clearButton = document.getElementById('customClearButton');
    const closeButton = document.getElementById('closeSearchResults');

    if (searchButton && searchInput) {
      searchButton.onclick = () => {
        this._performSearch(searchInput.value);
      };

      searchInput.onkeypress = (e: KeyboardEvent) => {
        if (e.key === 'Enter') {
          this._performSearch(searchInput.value);
        }
      };

      searchInput.oninput = () => {
        if (clearButton) {
          clearButton.style.display = searchInput.value ? 'flex' : 'none';
        }
      };
    }

    if (clearButton && searchInput) {
      clearButton.onclick = () => {
        searchInput.value = '';
        clearButton.style.display = 'none';
        this._closeSearch();
        searchInput.focus();
      };
    }

    if (closeButton) {
      closeButton.onclick = () => {
        this._closeSearch();
      };
    }

    // Close on click outside
    this._handleClickOutside = (event: MouseEvent) => {
      const searchContainer = document.querySelector(`.${styles.searchContainer}`);
      if (searchContainer && !searchContainer.contains(event.target as Node)) {
        this._closeSearch();
      }
    };
    document.addEventListener('click', this._handleClickOutside);
  }

  private _closeSearch(): void {
    const resultsElement = document.getElementById('customSearchResults');
    if (resultsElement) {
      resultsElement.style.display = 'none';
    }
  }

  private _performSearch(query: string): void {
    if (!query || query.trim() === '') return;

    const resultsElement = document.getElementById('customSearchResults');
    const resultsList = document.getElementById('resultsList');
    
    if (resultsElement && resultsList) {
      resultsElement.style.display = 'block';
      resultsList.innerHTML = '<div style="padding: 25px; text-align: center;"><i class="ms-Icon ms-Icon--ProgressRing" aria-hidden="true" style="font-size: 24px; margin-bottom: 10px; display: block;"></i> Searching...</div>';
    }

    const escapedQuery = query.replace(/'/g, "''");
    const searchUrl = `${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='${escapedQuery}'&selectproperties='Title,Path,Description'`;

    this.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (!response.ok) {
          throw new Error('Network response was not ok');
        }
        return response.json();
      })
      .then((data: ISearchResponse) => {
        this._displayResults(data);
      })
      .catch(error => {
        console.error("Search error:", error);
        if (resultsList) {
          resultsList.innerHTML = '<div style="padding: 15px; color: #d32f2f; text-align: center;">Error performing search. Please try again.</div>';
        }
      });
  }

  private _displayResults(data: ISearchResponse): void {
    const resultsElement = document.getElementById('customSearchResults');
    const resultsList = document.getElementById('resultsList');
    if (!resultsElement || !resultsList) return;

    try {
      const results = data.PrimaryQueryResult.RelevantResults.Table.Rows;

      if (!results || results.length === 0) {
        resultsList.innerHTML = '<div style="padding: 25px; text-align: center; color: #666;">No results found for your query.</div>';
        return;
      }

      let html = '';
      results.forEach((row: ISearchResultRow) => {
        const cells = row.Cells;
        const findCell = (key: string): { Key: string; Value: string } | undefined => cells.filter((c) => c.Key === key)[0];
        
        const titleCell = findCell('Title');
        const pathCell = findCell('Path');
        const descCell = findCell('Description');

        const title = titleCell ? titleCell.Value : 'No Title';
        const path = pathCell ? pathCell.Value : '#';
        const description = descCell ? descCell.Value : '';

        html += `
          <div class="${styles.resultItem}">
            <a href="${path}" target="_blank" data-interception="off">${escape(title)}</a>
            <div class="${styles.resultDesc}">${escape(description)}</div>
          </div>
        `;
      });

      resultsList.innerHTML = html;
    } catch (e) {
      console.error("Display error:", e);
      resultsList.innerHTML = '<div style="padding: 15px; color: #d32f2f; text-align: center;">Unable to process search results.</div>';
    }
  }

  protected onDispose(): void {
    if (this._handleClickOutside) {
        document.removeEventListener('click', this._handleClickOutside);
    }
    console.log('[HelloWorldApplicationCustomizer.onDispose] Disposed custom top placeholder and cleaned up events.');
  }
}
