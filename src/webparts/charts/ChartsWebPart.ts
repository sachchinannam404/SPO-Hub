import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import {
  IDataSourceConfig,
  IDisplayConfig,
  IChartConfig,
  DEFAULT_DATA_SOURCE,
  DEFAULT_DISPLAY
} from '../../models/Configuration';
import { ChartsContainer } from './components/ChartsContainer';

export interface IChartsWebPartProps {
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  chartConfig: IChartConfig;
}

export default class ChartsWebPart extends BaseClientSideWebPart<IChartsWebPartProps> {
  public render(): void {
    const element = React.createElement(ChartsContainer, {
      context: this.context,
      dataSource: this.properties.dataSource || { ...DEFAULT_DATA_SOURCE },
      display: this.properties.display || { ...DEFAULT_DISPLAY },
      chartConfig: this.properties.chartConfig || defaultChartConfig()
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onInit(): Promise<void> {
    if (!this.properties.dataSource) {
      this.properties.dataSource = {
        ...DEFAULT_DATA_SOURCE,
        itemLimit: 200,
        fieldMapping: {
          titleField: 'Title',
          categoryField: 'Category'
        }
      };
    }
    if (!this.properties.display) {
      this.properties.display = { ...DEFAULT_DISPLAY };
    }
    if (!this.properties.chartConfig) {
      this.properties.chartConfig = defaultChartConfig();
    }
    return super.onInit();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Chart data source and series' },
          groups: [
            {
              groupName: 'SharePoint List',
              groupFields: [
                PropertyPaneTextField('dataSource.listTitle', { label: 'List Title' }),
                PropertyPaneTextField('dataSource.filter', { label: 'OData Filter (optional)' }),
                PropertyPaneSlider('dataSource.itemLimit', {
                  label: 'Item Limit', min: 1, max: 500, step: 10, showValue: true
                })
              ]
            },
            {
              groupName: 'Chart Configuration',
              groupFields: [
                PropertyPaneTextField('display.title', { label: 'Chart Title' }),
                PropertyPaneDropdown('chartConfig.chartType', {
                  label: 'Chart Type',
                  options: [
                    { key: 'bar', text: 'Bar' },
                    { key: 'horizontalBar', text: 'Horizontal Bar' },
                    { key: 'line', text: 'Line' },
                    { key: 'area', text: 'Area' },
                    { key: 'pie', text: 'Pie' },
                    { key: 'doughnut', text: 'Doughnut' }
                  ]
                }),
                PropertyPaneTextField('chartConfig.categoryField', {
                  label: 'Category Field (X-Axis)',
                  description: 'Internal name of the category / label field'
                }),
                PropertyPaneTextField('chartConfig.series[0].field', {
                  label: 'Value Field (Y-Axis)',
                  description: 'Internal name of the numeric field'
                }),
                PropertyPaneDropdown('chartConfig.series[0].aggregation', {
                  label: 'Aggregation',
                  options: [
                    { key: 'sum', text: 'Sum' },
                    { key: 'count', text: 'Count' },
                    { key: 'avg', text: 'Average' },
                    { key: 'min', text: 'Min' },
                    { key: 'max', text: 'Max' }
                  ]
                }),
                PropertyPaneToggle('chartConfig.showLegend', { label: 'Show Legend' })
              ]
            }
          ]
        }
      ]
    };
  }
}

function defaultChartConfig(): IChartConfig {
  return {
    chartType: 'bar',
    categoryField: 'Category',
    series: [{ field: 'Value', aggregation: 'sum', label: 'Value' }],
    showLegend: true,
    showTitle: true
  };
}
