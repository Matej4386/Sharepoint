import {IFilterConfiguration} from './modules/datatypes/IFilterConfiguration';
import {SelectionMode} from 'office-ui-fabric-react/lib/DetailsList';
import {SPHttpClient} from '@microsoft/sp-http';

export interface IMSptablev2Props {
  /**
   * from context not to pass whole context if it is possible
   */
  currentCultureName: string;
  spHttpClient: SPHttpClient; 
  webAbsoluteUrl: string;
  /**
   * console.log info
   */
  debug: boolean;
  /**
   * List ID to display
   */
  listToDisplayId: string;
  /**
   * Filters configuration
   */
  filterConfiguration: IFilterConfiguration[];
  /**
   * String with internal column names separated with ,
   */
  searchConfiguration: string;
  /**
   * selection mode - none, single, multiple
   */
  selectionMode: SelectionMode;
  /**
   * function to render info (errors, success) in parent element
   */
  renderInfo: (error: boolean, message: string) => void;
}
