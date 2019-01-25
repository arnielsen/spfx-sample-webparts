import * as React from 'react';
import styles from './DelveBlogPosts.module.scss';
import { IDelveBlogPostsProps } from './IDelveBlogPostsProps';
import { IDelveBlogPostsState } from './IDelveBlogPostsState';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/components/Spinner';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  sp,
  SearchQuery,
  SearchResults,
  SearchQueryBuilder
} from '@pnp/sp';
import Moment from 'react-moment';

export default class DelveBlogPosts extends React.Component<IDelveBlogPostsProps, IDelveBlogPostsState, {}> {
  
  /**
   * Create an instance of DelveBlogPosts
   */
  public constructor(props: IDelveBlogPostsProps) {
    super(props);

    this.state = {
      items: [],
      loading: false
    };
  }

  /**
   * componentDidMount lifecycle hook
   */
  public componentDidMount() {
    this.renderContents();
  }

  /** 
   * componentDidUpdate lifecycle hook
   */
  public componentDidUpdate() {
    //this.renderContents();
  }

  /**
   * Retrieve data and render contents
   */
  private renderContents() {
    // Update state
    this.setState({
      loading: true
    });

    // Check environment type
    // todo: implement mock date?

    // Construct query
    let queryText = "filetype:PointPub";

    // Handle sorting
    const sortList = {
      Property:"Created",
      Direction:1
    };
    
    // Call Search API
    const builder = SearchQueryBuilder().text(queryText).rowLimit(this.props.rowLimit | 3).sortList(sortList);
    sp.searchWithCaching(builder).then(r => {
      console.log(r.ElapsedTime);
      console.log(r.RowCount);
      console.log(r.PrimarySearchResults);

      this.setState({
        items: r.PrimarySearchResults,
        loading: false
      });
    }).catch(c => {
      // error handling
    });
  }
  
  public render(): React.ReactElement<IDelveBlogPostsProps> {
    return (
      <div>
        {
          this.state.loading ? 
          (
            <Spinner size={SpinnerSize.large} label="Loading blogs..." />
          ) : (
            this.state.items.length === 0 ? 
            (
              <div>
                Did not find any blogs.. Why don't you create one?
              </div>
            ) : (
              this.state.items.map(item => {
                return (
                  <div>
                    <div>{item.Title}</div>
                    <div>{item.Author}</div>
                    <div>{item.Write}</div>
                    <div>{item.Path}</div>
                    <div>{item.PictureThumbnailURL}</div>
                  </div>
                );
              })
            )
          )
        }
      </div>
    );
  }
}
