//import ISPList interface to this file from PnPjsHelloWorldWebPart.ts
import { ISPList } from './PnPjsHelloWorldWebPart';
 
//Declare ISPList array and returns the array of items whenever MockHttpClient.get() method called
export default class MockHttpClient {
 
    private static _items: ISPList[] = [{ Title: 'Mock List', Id: '1' }];
 
    public static get(restUrl: string, options?: any): Promise<ISPList[]> {
      return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}