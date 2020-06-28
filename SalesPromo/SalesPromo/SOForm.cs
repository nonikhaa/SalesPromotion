using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesPromo
{
    public class SOForm
    {
        private string cardCode;
        private string postingDate;
        private string uniqueID;
        private string shipTo;

        private bool itemCodeState;
        private bool qtyState;
        private bool uomState;
        private bool discAddonState;
        private bool discPrcntState;
        private bool fixDiscState;
        private bool prdDiscState;
        private bool bonusItem;

        public string CardCode
        {
            get { return cardCode; }
            set { cardCode = value; }
        }

        public string PostingDate
        {
            get { return postingDate; }
            set { postingDate = value; }
        }

        public string UniqueID
        {
            get { return uniqueID; }
            set { uniqueID = value; }
        }

        public string ShipTo
        {
            get { return shipTo; }
            set { shipTo = value; }
        }

        public bool ItemCodeState
        {
            get { return itemCodeState; }
            set { itemCodeState = value; }
        }

        public bool QtyState
        {
            get { return qtyState; }
            set { qtyState = value; }
        }

        public bool UomState
        {
            get { return uomState; }
            set { uomState = value; }
        }

        public bool DiscAddonState
        {
            get { return discAddonState; }
            set { discAddonState = value; }
        }

        public bool DiscPrcntState
        {
            get { return discPrcntState; }
            set { discPrcntState = value; }
        }

        public bool FixDiscState
        {
            get { return fixDiscState; }
            set { fixDiscState = value; }
        }

        public bool PrdDiscState
        {
            get { return prdDiscState; }
            set { prdDiscState = value; }
        }

        public bool BonusItem
        {
            get { return bonusItem; }
            set { bonusItem = value; }
        }

        public void GetColumnState(bool itemCode, bool qty, bool uom, bool discAddon, bool fixDisc, bool prdDisc, bool bonusItem)
        {
            ItemCodeState = itemCode;
            QtyState = qty;
            UomState = uom;
            DiscAddonState = discAddon;
            FixDiscState = fixDisc;
            PrdDiscState = prdDisc;
            BonusItem = bonusItem;
        }
    }

    public class MatrixSo
    {
        public string ItemCode;
        public double Quantity;
        public string Area;
        public string AreaType;
        public double SapDisc;
        public string Address;
        public string FlagBonus;

    }


    public class OutputDiscQuery
    {
        public string DiscountType;
        public string CustomerType;
        public string CustomerCode;
        public string ItemCode;
        public string Area;

        public string FixDiscCode;
        public double FixDisc;

        public string PrdDiscCode;
        public double PrcntDisc;
        public double PrcntMinQty;

        public string BXGYItemCd;
        public double BXGYMinQty;
        public string BXGYItemBns;
        public double BXGYQtyFree;
        public string Kelipatan;

    }
}
