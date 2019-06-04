import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GeoMapWebPart.module.scss';
import * as strings from 'GeoMapWebPartStrings';
import * as d3 from "d3";
import * as topojson from 'topojson';
import { geoPath } from "d3-geo";
import { geoRobinson } from "d3-geo-projection";
import * as d3Tip from "d3-tip";
require('../../assets/d3-tip.css');

export interface IGeoMapWebPartProps {
  description: string;
}

export default class GeoMapWebPart extends BaseClientSideWebPart<IGeoMapWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.geoMap}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
            <div class="map"></div>
            </div>
          </div>
        </div>
      </div>`;
    const format = d3.format(',');

    // Set tooltips
    const tip = d3Tip.default()
      .attr('class', 'd3-tip')
      .offset([-10, 0])
      .html(d => `<strong>Country: </strong><span class='details'>${d.properties.name}<br></span><strong>Population: </strong><span class='details'>${format(d.population)}</span>`);

    const margin = { top: 0, right: 400, bottom: 0, left: 0 };
    const width = 960 - margin.left - margin.right;
    const height = 500 - margin.top - margin.bottom;

    const color = d3.scaleThreshold()
      .domain([
        10000,
        100000,
        500000,
        1000000,
        5000000,
        10000000,
        50000000,
        100000000,
        500000000,
        1500000000
      ])
      .range([
        'rgb(247,251,255)',
        'rgb(222,235,247)',
        'rgb(198,219,239)',
        'rgb(158,202,225)',
        'rgb(107,174,214)',
        'rgb(66,146,198)',
        'rgb(33,113,181)',
        'rgb(8,81,156)',
        'rgb(8,48,107)',
        'rgb(3,19,43)'
      ]);

    const svg = d3.select(document.querySelector('.map'))
      .append('svg')
      .attr('width', width)
      .attr('height', height)
      .append('g')
      .attr('class', 'map');

    const projection = geoRobinson()
      .scale(100)
      .rotate([352, 0, 0])
      .translate([width / 2, height / 2]);

    const path = geoPath().projection(projection);

    svg.call(tip);
    // upload below files to assets library/cdn and provide the path
    Promise.all([
      d3.json('world_countries.txt'),
      d3.tsv('world_population.tsv')
    ]).then(
      d => ready(null, d[0], d[1])
    );

    function ready(error, data, population) {
      const populationById = {};

      population.forEach(d => { populationById[d.id] = +d.population; });
      data.features.forEach(d => { d.population = populationById[d.id]; });

      svg.append('g')
        .attr('class', 'countries')
        .selectAll('path')
        .data(data.features)
        .enter().append('path')
        .attr('d', path)
        .style('fill', d => color(populationById[d.id]))
        .style('stroke', 'white')
        .style('opacity', 0.8)
        .style('stroke-width', 0.3)
        // tooltips
        .on('mouseover', function (d) {
          tip.show(d, this);
          d3.select(this)
            .style('opacity', 1)
            .style('stroke-width', 3);
        })
        .on('mouseout', function (d) {        
         tip.hide(d, this);
          d3.select(this)
            .style('opacity', 0.8)
            .style('stroke-width', 0.3);
        });

      svg.append('path')
        .datum(topojson.mesh(data.features, (a, b) => a.id !== b.id))
        .attr('class', 'names')
        .attr('d', path);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
