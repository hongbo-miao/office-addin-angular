import { AngularOfficeAddinPage } from './app.po';

describe('angular-office-addin App', () => {
  let page: AngularOfficeAddinPage;

  beforeEach(() => {
    page = new AngularOfficeAddinPage();
  });

  it('should display message saying app works', () => {
    page.navigateTo();
    expect(page.getParagraphText()).toEqual('app works!');
  });
});
