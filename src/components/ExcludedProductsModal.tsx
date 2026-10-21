import { Modal } from "@/components/Modal";
import type { ExcludedProduct } from "@/lib/types";

function Status({ present }: { present: boolean }) {
  return <span className={`metadata-status ${present ? "is-present" : "is-missing"}`}>{present ? "Yes / हाँ" : "No / नहीं"}</span>;
}

export function ExcludedProductsModal({ products, onClose }: { products: ExcludedProduct[]; onClose: () => void }) {
  return (
    <Modal title="Excluded product details / छोड़े गए प्रोडक्ट की जानकारी" onClose={onClose} wide>
      <p className="modal-help">These products were not added to PS or LI because required information is missing. / आवश्यक जानकारी न होने के कारण ये प्रोडक्ट PS या LI में शामिल नहीं किए गए।</p>
      <div className="sheet-scroll" tabIndex={0}>
        <table className="sheet-table metadata-table">
          <thead><tr><th scope="col">Product name<br /><span lang="hi">प्रोडक्ट नाम</span></th><th scope="col">TRAN_TYPE</th><th scope="col">ORDER_TYPE</th><th scope="col">Product code<br /><span lang="hi">प्रोडक्ट कोड</span></th></tr></thead>
          <tbody>
            {products.map((product) => (
              <tr key={product.columnLetter}>
                <td><strong>{product.name}</strong><br /><small>Column / कॉलम {product.columnLetter}</small></td>
                <td><Status present={product.hasTranType} /></td>
                <td><Status present={product.hasOrderType} /></td>
                <td><Status present={product.hasItemCode} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}
