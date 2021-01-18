# m-sptablev-2

## Summary

Sharepoint react custom table with server filtering

## Prerequisites

change ShimmeredDetailsLits.d.ts:
/// <reference types="react" />
import * as React from 'react';
import { BaseComponent } from '../../Utilities';
import { IDetailsListProps } from './DetailsList.types';
export interface IShimmeredDetailsListProps extends IDetailsListProps {
    shimmerLines?: number;
    onRenderCustomPlaceholder?: (rowProps: IDetailsRowProps, index?: number, defaultRender?: (props: IDetailsRowProps) => React.ReactNode) => React.ReactNode;
}
export declare class ShimmeredDetailsList extends BaseComponent<IShimmeredDetailsListProps, {}> {
    private _shimmerItems;
    constructor(props: IShimmeredDetailsListProps);
    render(): JSX.Element;
    private _onRenderShimmerPlaceholder;
    private _renderDefaultShimmerPlaceholder;
}

change ShimmeredDetailsLits.js:
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = require("react");
var Utilities_1 = require("../../Utilities");
var interfaces_1 = require("../../utilities/selection/interfaces");
var DetailsList_types_1 = require("./DetailsList.types");
var DetailsList_1 = require("./DetailsList");
var Shimmer_1 = require("../Shimmer");
var rowStyles = require("./DetailsRow.scss");
var listStyles = require("./DetailsList.scss");
var SHIMMER_INITIAL_ITEMS = 10;
var DEFAULT_SHIMMER_HEIGHT = 7;
var SHIMMER_LINE_VS_CELL_WIDTH_RATIO = 0.95;
// This values are matching values from ./DetailsRow.css
var DEFAULT_SIDE_PADDING = 8;
var DEFAULT_EXTRA_SIDE_PADDING = 24;
var DEFAULT_ROW_HEIGHT = 42;
var COMPACT_ROW_HEIGHT = 32;
var ShimmeredDetailsList = /** @class */ (function (_super) {
    tslib_1.__extends(ShimmeredDetailsList, _super);
    function ShimmeredDetailsList(props) {
        var _this = _super.call(this, props) || this;
        _this._onRenderShimmerPlaceholder = function (index, rowProps) {
            var _a = _this.props, onRenderCustomPlaceholder = _a.onRenderCustomPlaceholder, compact = _a.compact;
            var selectionMode = rowProps.selectionMode, checkboxVisibility = rowProps.checkboxVisibility;
            var showCheckbox = selectionMode !== interfaces_1.SelectionMode.none && checkboxVisibility !== DetailsList_types_1.CheckboxVisibility.hidden;
            var placeholderElements = onRenderCustomPlaceholder
                ? onRenderCustomPlaceholder(rowProps, index, _this._renderDefaultShimmerPlaceholder)
                : _this._renderDefaultShimmerPlaceholder(rowProps);
            return (React.createElement("div", { className: Utilities_1.css(showCheckbox && rowStyles.shimmerLeftBorder, !compact && rowStyles.shimmerBottomBorder) },
                React.createElement(Shimmer_1.Shimmer, { customElementsGroup: placeholderElements })));
        };
        _this._renderDefaultShimmerPlaceholder = function (rowProps) {
            var columns = rowProps.columns, compact = rowProps.compact;
            var shimmerElementsRow = [];
            var gapHeight = compact ? COMPACT_ROW_HEIGHT : DEFAULT_ROW_HEIGHT;
            columns.map(function (column, columnIdx) {
                var shimmerElements = [];
                var groupWidth = DEFAULT_SIDE_PADDING * 2 +
                    column.calculatedWidth +
                    (column.isPadded ? DEFAULT_EXTRA_SIDE_PADDING : 0);
                shimmerElements.push({
                    type: Shimmer_1.ShimmerElementType.gap,
                    width: DEFAULT_SIDE_PADDING,
                    height: gapHeight
                });
                if (column.isIconOnly) {
                    shimmerElements.push({
                        type: Shimmer_1.ShimmerElementType.line,
                        width: column.calculatedWidth,
                        height: column.calculatedWidth
                    });
                    shimmerElements.push({
                        type: Shimmer_1.ShimmerElementType.gap,
                        width: DEFAULT_SIDE_PADDING,
                        height: gapHeight
                    });
                }
                else {
                    shimmerElements.push({
                        type: Shimmer_1.ShimmerElementType.line,
                        width: column.calculatedWidth * SHIMMER_LINE_VS_CELL_WIDTH_RATIO,
                        height: DEFAULT_SHIMMER_HEIGHT
                    });
                    shimmerElements.push({
                        type: Shimmer_1.ShimmerElementType.gap,
                        width: DEFAULT_SIDE_PADDING +
                            (column.calculatedWidth - column.calculatedWidth * SHIMMER_LINE_VS_CELL_WIDTH_RATIO) +
                            (column.isPadded ? DEFAULT_EXTRA_SIDE_PADDING : 0),
                        height: gapHeight
                    });
                }
                shimmerElementsRow.push(React.createElement(Shimmer_1.ShimmerElementsGroup, { key: columnIdx, width: groupWidth + "px", shimmerElements: shimmerElements }));
            });
            // When resizing the window from narrow to wider, we need to cover the exposed Shimmer wave until the column resizing logic is done.
            shimmerElementsRow.push(React.createElement(Shimmer_1.ShimmerElementsGroup, { key: 'endGap', width: '100%', shimmerElements: [{ type: Shimmer_1.ShimmerElementType.gap, width: '100%', height: gapHeight }] }));
            return React.createElement("div", { style: { display: 'flex' } }, shimmerElementsRow);
        };
        _this._shimmerItems = props.shimmerLines ? new Array(props.shimmerLines) : new Array(SHIMMER_INITIAL_ITEMS);
        return _this;
    }
    ShimmeredDetailsList.prototype.render = function () {
        var _a = this.props, items = _a.items, listProps = _a.listProps;
        var _b = this.props, shimmerLines = _b.shimmerLines, onRenderCustomPlaceholder = _b.onRenderCustomPlaceholder, enableShimmer = _b.enableShimmer, detailsListProps = tslib_1.__rest(_b, ["shimmerLines", "onRenderCustomPlaceholder", "enableShimmer"]);
        // Adds to the optional listProp classname a fading out overlay classname only when shimmer enabled.
        var shimmeredListClassname = Utilities_1.css(listProps && listProps.className, enableShimmer && listStyles.shimmerFadeOut);
        var newListProps = tslib_1.__assign({}, listProps, { className: shimmeredListClassname });
        return (React.createElement(DetailsList_1.DetailsList, tslib_1.__assign({}, detailsListProps, { items: enableShimmer ? this._shimmerItems : items, onRenderMissingItem: this._onRenderShimmerPlaceholder, listProps: newListProps })));
    };
    return ShimmeredDetailsList;
}(Utilities_1.BaseComponent));
exports.ShimmeredDetailsList = ShimmeredDetailsList;
//# sourceMappingURL=ShimmeredDetailsList.js.map

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

