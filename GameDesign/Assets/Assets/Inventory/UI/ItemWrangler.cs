using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;
using UnityEngine.EventSystems;
using UnityEngine.InputSystem;
using System.ComponentModel.Design;

public class ItemWrangler : MonoBehaviour, IPointerEnterHandler {
    public enum State {
        Picking,
        Dragging    
    }

    [SerializeField] State _state;

    [SerializeField] InputActionAsset _inputMap;
    InputAction _pointer;
    InputAction _click;

    [SerializeField] Dragger _dragger;



    [SerializeField] AnimationCurve _hoverJuice;
    
    ItemTile _active;
    ItemTile _dragging;
    List<ItemTile> _deselecting;
    List<RaycastResult> _hits;

    float _activeTime; 
    float _draggingTime;   

    void Start() {
        _state = State.Picking;
        _deselecting = new();
        _hits = new();
        _pointer = _inputMap.FindAction("Point");        
        _click = _inputMap.FindAction("Click");
        _click.performed += OnClick;
    }

    void OnDestroy() {
        _click.performed -= OnClick;
    }

    void OnClick(InputAction.CallbackContext context) {
        if (_state == State.Picking && _active != null) {
            _dragging = _active;
            _dragger.Select(_active.item.sprite);
            Vector2 pointerPos = _pointer.ReadValue<Vector2>();
            RectTransformUtility.ScreenPointToLocalPointInRectangle((RectTransform)_dragger.transform.parent, pointerPos, null, out var cursorPos);
            _dragger.transform.localPosition = cursorPos;
            _dragger.gameObject.SetActive(true);
            _state = State.Dragging;
            _draggingTime = 0f;
        }
    }
    
    void Update() {        
        if (_state == State.Dragging && !_click.IsPressed()) {
            // TODO: Drop logic.
            _draggingTime = 0f;
            _dragging = null;
            _state = State.Picking;            
        }

        _draggingTime += Time.deltaTime;

        if (_state != State.Dragging && _dragger.gameObject.activeSelf) {
            if (_dragger.DropAndDisable(_draggingTime)) _dragger.gameObject.SetActive(false);
        }

        var pointerPos =  _pointer.ReadValue<Vector2>();
        if (_active != null) {

            Vector2 cursorPos;
            RectTransformUtility.ScreenPointToLocalPointInRectangle((RectTransform)_active.transform, pointerPos, null, out cursorPos);
            const float hoverRange = 100f;
            //debugInfo = cursorPos.ToString();
            cursorPos.y *= 0.5f;
            if (_state == State.Picking) {
                if (cursorPos.sqrMagnitude > hoverRange*hoverRange) {
                    debugInfo = "deselecting";
                    Deselect();
                    _active = null;
                } else {
                    _activeTime += Time.unscaledDeltaTime;
                    _active.SetHighlight(_hoverJuice.Evaluate(_activeTime));
                }
            } else if (_state == State.Dragging) {
                _dragger.Highlight(_draggingTime);
                RectTransformUtility.ScreenPointToLocalPointInRectangle((RectTransform)_dragger.transform.parent, pointerPos, null, out cursorPos);
                _dragger.transform.localPosition = cursorPos;
            }
        }

        float fadeMultiplier = Mathf.Pow(0.01f, Time.unscaledDeltaTime);
        foreach (var tile in _deselecting) {
            tile.SetHighlight(tile.highlight * fadeMultiplier);
        }

        _deselecting.RemoveAll(t => t.highlight < 0.001f);
    }

    void Deselect() {
        _deselecting.Add(_active);
    }

    int eventCount = 0;
    [TextArea(6,6)]
    public string debugInfo;

    public void OnPointerEnter(PointerEventData eventData) {
        
        eventCount++;
        debugInfo = eventCount.ToString();
        foreach(var hover in eventData.hovered) {
            var parent = hover.transform.parent;
            debugInfo += $"\n{(parent == null ? "null" : parent.name)}->{hover.name}";
        }

        if (eventData.hovered.Count == 0) return;
        

        var tile = eventData.hovered[0].GetComponentInParent<ItemTile>();
        if (tile == null) {
            debugInfo += "\n null tile";
            return;
        }

        _deselecting.Remove(tile);

        if (_active != null && _active != tile) {
            Deselect();
        }

        if (tile != _active)
            _activeTime = 0f;

        _active = tile;
        
        debugInfo += "\n selecting";
    }
}
