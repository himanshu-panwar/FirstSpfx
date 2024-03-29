import { ISPList } from './HelloWorldWebPart';

export default class MockHttpClient{
    private static _items: ISPList[] = [{ Title: 'Mock List 1', Id:'1' },
                                        { Title: 'Mock List 2', Id:'2' }];

    public static get():Promise<ISPList[]> {
        return new Promise<ISPList[]>((resolve) => {
            resolve(MockHttpClient._items);
        });
    }
}