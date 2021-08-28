// https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/926f1c2ba0598575e23dfd8cdd8b79fa3a3d19ff/samples/msal-angular-v2-samples/angular10-browser-sample/src/app/msal/msal.interceptor.ts
import {
  HttpRequest,
  HttpHandler,
  HttpEvent,
  HttpInterceptor
} from '@angular/common/http';
import { Observable, from, EMPTY } from 'rxjs';
import { switchMap, catchError } from 'rxjs/operators';
import { MsalService } from '@azure/msal-angular';
import { Minimatch } from "minimatch";
import { PopupRequest, RedirectRequest, AuthenticationResult, InteractionType } from "@azure/msal-browser";
import { Injectable, Inject } from '@angular/core';
import { MSAL_INTERCEPTOR_CONFIG } from '@azure/msal-angular';

export type MsalInterceptorConfig = {
  interactionType: InteractionType.Popup | InteractionType.Redirect;
  protectedResourceMap: Map<string, Array<string>>;
  authRequest?: PopupRequest | RedirectRequest;
};

@Injectable()
export class AzureMsalInterceptor implements HttpInterceptor {
  urlsToUse: Array<string>;
  constructor(
      @Inject(MSAL_INTERCEPTOR_CONFIG) private msalInterceptorConfig: MsalInterceptorConfig,
      private authService: MsalService
  ) {
    this.urlsToUse= [
      'azure/.+',
     // 'myController1/myAction3'
    ];
  }

  intercept(req: HttpRequest<any>, next: HttpHandler): Observable<HttpEvent<any>> {
      const scopes = this.getScopesForEndpoint(req.url);
      const account = this.authService.instance.getActiveAccount() || this.authService.instance.getAllAccounts()[0];

      if (!scopes || scopes.length === 0) {
          return next.handle(req);
      }

      // Note: For MSA accounts, include openid scope when calling acquireTokenSilent to return idToken
      return this.authService.acquireTokenSilent({scopes, account})
          .pipe(
              catchError(() => {
                  if (this.msalInterceptorConfig.interactionType === InteractionType.Popup) {
                      return this.authService.acquireTokenPopup({...this.msalInterceptorConfig.authRequest, scopes});
                  }
                  const redirectStartPage = window.location.href;
                  this.authService.acquireTokenRedirect({...this.msalInterceptorConfig.authRequest, scopes, redirectStartPage});
                  return EMPTY;
              }),
              switchMap((result: AuthenticationResult) => {
                  if (this.isValidRequestForInterceptor(req.url)) {
                    const headers = req.headers
                      .set('Authorization', `Bearer ${result.accessToken}`);

                    const requestClone = req.clone({headers});
                    return next.handle(requestClone);
                  }
                  return next.handle(req);
              })
          );

  }

  // https://stackoverflow.com/questions/55522320/angular-interceptor-exclude-specific-urls
  private isValidRequestForInterceptor(requestUrl: string): boolean {
    let positionIndicator: string = 'api/';
    let position = requestUrl.indexOf(positionIndicator);
    if (position > 0) {
      let destination: string = requestUrl.substr(position + positionIndicator.length);
      for (let address of this.urlsToUse) {
        if (new RegExp(address).test(destination)) {
          return true;
        }
      }
    }
    return false;
  }

  private getScopesForEndpoint(endpoint: string): Array<string>|undefined {
      const protectedResourcesArray = Array.from(this.msalInterceptorConfig.protectedResourceMap.keys());
      const keyMatchesEndpointArray = protectedResourcesArray.filter(key => {
          const minimatch = new Minimatch(key);
          return minimatch.match(endpoint) || endpoint.indexOf(key) > -1;
      });

      // process all protected resources and send the first matched resource
      if (keyMatchesEndpointArray.length > 0) {
          const keyForEndpoint = keyMatchesEndpointArray[0];
          if (keyForEndpoint) {
              return this.msalInterceptorConfig.protectedResourceMap.get(keyForEndpoint);
          }
      }

      return undefined;
  }

}