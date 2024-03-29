
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SPHttpClient, SPHttpClientResponse, MSGraphClient } from "@microsoft/sp-http";
import * as moment from 'moment';

export class SPService {
  private graphClient: MSGraphClient = null;
  private  birthdayListTitle: string = "Birthdays";
  constructor(private _context: WebPartContext | ApplicationCustomizerContext) {

  }

  private fixUserArray(arr: any){
    return arr.PrimaryQueryResult.RelevantResults.Table.Rows.map((row) => {
      return row.Cells.reduce((res, cell) => {
          res[cell.Key] = cell.Value;
          return res;
      }, {});
    });
  }

  public async GetUsers (context: WebPartContext){
    try {
      let data1 = await context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/search/query?querytext='*'&sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'&selectproperties='Title,AccountName,Path,WorkEmail,RefinableDate00,OfficeNumber,Department,JobTitle,WorkPhone'&rowlimit=500&startrow=1`, SPHttpClient.configurations.v1);
      let data2 = await context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/search/query?querytext='*'&sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'&selectproperties='Title,AccountName,Path,WorkEmail,RefinableDate00,OfficeNumber,Department,JobTitle,WorkPhone'&rowlimit=500&startrow=501`, SPHttpClient.configurations.v1);
      let data3 = await context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/search/query?querytext='*'&sourceid='B09A7990-05EA-4AF9-81EF-EDFAB16C4E31'&selectproperties='Title,AccountName,Path,WorkEmail,RefinableDate00,OfficeNumber,Department,JobTitle,WorkPhone'&rowlimit=500&startrow=1001`, SPHttpClient.configurations.v1);
      let jsonData1 = await data1.json();
      let jsonData2 = await data2.json();
      let jsonData3 = await data3.json();

      let results1 = this.fixUserArray(jsonData1);
      let results2 = this.fixUserArray(jsonData2);
      let results3 = this.fixUserArray(jsonData3);
  
      let fullCall = results1.concat(results2).concat(results3);
      let fullCallSorted = fullCall.filter(a => a.WorkEmail != "" && a.WorkEmail != null);

      return fullCallSorted;

    } catch (error) {
      console.log(error);
      return(null);
    }

  }
  // Get Profiles
  public async getPBirthdays(upcommingDays: number): Promise<any[]> {
    let _results, _today: string, _month: string, _day: number;
    let _filter: string, _countdays: number, _f:number, _nextYearStart: string;
    let  _FinalDate: string;
    try {
      _results = null;
      _today = '2000-' + moment().format('MM-DD');
      _month = moment().format('MM');
      _day = parseInt(moment().format('DD'));
      _filter = "fields/Birthday ge '" + _today + "'";
      // If we are in Dezember we have to look if there are birthday in January
      // we have to build a condition to select birthday in January based on number of upcommingDays
      // we can not use the year for teste , the year is always 2000.
      console.log(_month);
      if (_month === '12') {
        _countdays = _day + upcommingDays;
        _f = 0;
        _nextYearStart = '2000-01-01';
        _FinalDate = '2000-01-';
        if ((_countdays) > 31) {
          _f = _countdays - 31;
          _FinalDate = _FinalDate + _f;
          _filter = "fields/Birthday ge '" + _today + "' or (fields/Birthday ge '" + _nextYearStart + "' and fields/Birthday le '" + _FinalDate + "')";
        }
      }
      this.graphClient = await this._context.msGraphClientFactory.getClient();
      _results = await this.graphClient.api(`sites/root/lists('${this.birthdayListTitle}')/items?orderby=Fields/Birthday`)
        .version('v1.0')
        .expand('fields')
        .top(upcommingDays)
        .filter(_filter)
        .get();

        return _results.value;

    } catch (error) {
      console.dir(error);
      return Promise.reject(error);
    }
  }
}
export default SPService;
